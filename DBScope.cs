// by Brian Lalonde http://webcoder.info/downloads/
// This work is licensed under the Creative Commons Attribution-Share Alike 3.0 License. 
// To view a copy of this license, visit http://creativecommons.org/licenses/by-sa/3.0/ 
// or send a letter to Creative Commons, 543 Howard Street, 5th Floor, San Francisco, California, 94105, USA.

using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.Common;
using System.Diagnostics;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace DBScope
{
	public partial class Scope : Form
	{
		readonly string[] StatsQuery;

		private class DistinguishedConnection : IDisposable
		{
			public DbConnection Connection = null;

			public DataTable Schema = null;

			public DataTable Info = null;

			public DataTable Types = null;

			public DataTable ReservedWords = null;

			public DistinguishedConnection() { }

			public DistinguishedConnection(DbProviderFactory factory, ConnectionStringSettings connstr)
			{
				Connection = factory.CreateConnection();
				Connection.ConnectionString = connstr.ConnectionString;
				Connection.Open();
				Schema = Connection.GetSchema();
				Info = Connection.GetSchema("DataSourceInformation");
				Types = Connection.GetSchema("DataTypes");
				ReservedWords = Connection.GetSchema("ReservedWords");
			}

			public bool Has(string collection)
			{
				return Schema.Select(String.Format("CollectionName='{0}'",
					collection.Replace("'", "''"))).Length > 0;
			}

			public void Dispose()
			{
				if (Connection != null) Connection.Dispose();
				if (Schema != null) Schema.Dispose();
			}
		}

		private class Context : DistinguishedConnection
		{
			public DbProviderFactory Factory = null;

			public string TableName = null;

			public string[] TableFilter = null;

			public int Depth = 0;

			public Context(TreeNode node)
			{
				string connstr = null, database = null;
				for (; node != null; node = node.Parent, Depth++)
				{
					if (node.Tag is DbProviderFactory) { Factory = (DbProviderFactory)node.Tag; break; }
					if (node.Tag is string[]) { TableName = node.Text; TableFilter = (string[])node.Tag; continue; }
					if (node.Tag is string) { database = (string)node.Tag; continue; }
					if (node.Tag is DistinguishedConnection)
					{
						Connection = ((DistinguishedConnection)node.Tag).Connection;
						Schema = ((DistinguishedConnection)node.Tag).Schema;
						Info = ((DistinguishedConnection)node.Tag).Info;
						Types = ((DistinguishedConnection)node.Tag).Types;
						ReservedWords = ((DistinguishedConnection)node.Tag).ReservedWords;
					}
					else if (node.Tag is ConnectionStringSettings)
						connstr = ((ConnectionStringSettings)node.Tag).ConnectionString;
				}
				if (connstr != null)
				{
					Connection = Factory.CreateConnection();
					Connection.ConnectionString = connstr;
					Connection.Open();
					if (!String.IsNullOrEmpty(database) && (Connection.Database != database))
						Connection.ChangeDatabase(database);
					Schema = Connection.GetSchema();
					Info = Connection.GetSchema("DataSourceInformation");
					Types = Connection.GetSchema("DataTypes");
					ReservedWords = Connection.GetSchema("ReservedWords");
				}
			}

			public DataTable GetSchema(string collection)
			{
				return Connection.GetSchema(collection);
			}

			public DataTable GetColumns()
			{
				return Connection.GetSchema("Columns", TableFilter);
			}

			public string Database
			{
				get { return Connection.Database; }
				set { Connection.ChangeDatabase(value); }
			}

			public DbCommand CreateCommand()
			{
				return Connection.CreateCommand();
			}
		}

		public Scope()
		{
			StringCollection queries = new StringCollection();
			for (int i = 0; ConfigurationManager.AppSettings["StatsQuery." + i.ToString()] != null; i++)
				queries.Add(ConfigurationManager.AppSettings["StatsQuery." + i.ToString()]);
			StatsQuery = new string[queries.Count];
			queries.CopyTo(StatsQuery, 0);
			InitializeComponent();
		}

		private void Scope_Load(object sender, EventArgs e)
		{
			DataTree.Nodes.Add(new TreeNode("About"));
			DataTree.Nodes[0].ForeColor = Color.Blue;
			DataTable factories = DbProviderFactories.GetFactoryClasses();
			foreach (DataRow factoryinfo in factories.Rows)
			{
				DbProviderFactory factory = null;
				try { factory = DbProviderFactories.GetFactory(factoryinfo); }
				catch (ConfigurationErrorsException cfgex)
				{
					Trace.TraceError(factoryinfo[0].ToString() + "\n" + factoryinfo[1].ToString() + "\n" +
						factoryinfo[2].ToString() + "\n" + factoryinfo[3].ToString() + "\n" + cfgex.ToString());
					continue;
				}
				TreeNode factorynode = new TreeNode((string)factoryinfo[0]);
				string typename = factory.GetType().FullName;
				factorynode.Tag = factory;
				DataTree.Nodes.Add(factorynode);
				foreach (ConnectionStringSettings cs in ConfigurationManager.ConnectionStrings)
				{
					if (cs.ProviderName != (string)factoryinfo[2]) continue;
					TreeNode connode = new TreeNode(cs.Name);
					connode.Tag = cs;
					factorynode.Nodes.Add(connode);
					connode.Nodes.Add(new TreeNode());
				}
				RegistryKey dsns;
				if (typename == "System.Data.Odbc.OdbcFactory")
				{
					dsns = Registry.LocalMachine.OpenSubKey(@"Software\ODBC\ODBC.INI\ODBC Data Sources");
					if (dsns != null)
						foreach (string value in dsns.GetValueNames())
						{
							TreeNode dsn = new TreeNode(value);
							dsn.Tag = new ConnectionStringSettings(value, "dsn=" + value, typename);
							dsn.Nodes.Add(new TreeNode());
							factorynode.Nodes.Add(dsn);
						}
					dsns = Registry.CurrentUser.OpenSubKey(@"Software\ODBC\ODBC.INI\ODBC Data Sources");
					if (dsns != null)
						foreach (string value in dsns.GetValueNames())
						{
							TreeNode dsn = new TreeNode(value);
							dsn.Tag = new ConnectionStringSettings(value, "dsn=" + value, typename);
							dsn.Nodes.Add(new TreeNode());
							factorynode.Nodes.Add(dsn);
						}
				}
				dsns = Registry.CurrentUser.OpenSubKey(@"Software\webcoder\DBScope\" + typename);
				if (dsns != null)
					foreach (string value in dsns.GetValueNames())
					{
						TreeNode dsn = new TreeNode(value);
						dsn.Tag = new ConnectionStringSettings(value, (string)dsns.GetValue(value), typename);
						dsn.Nodes.Add(new TreeNode());
						factorynode.Nodes.Add(dsn);
					}
				factorynode.Expand();
			}
		}

		private void DataTree_BeforeExpand(object sender, TreeViewCancelEventArgs e)
		{
			if ((e.Node.Nodes.Count != 1) || (e.Node.Nodes[0].Text != "")) return;
			Context context = null;
			try { context = new Context(e.Node); }
			catch (DbException dbex)
			{
				Trace.TraceError(dbex.ToString());
				MessageBox.Show(dbex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				e.Node.Nodes.Clear();
			}
			switch (context.Depth)
			{
				case 1: GetDatabases(e.Node, context); break;
				case 2: GetTables(e.Node, context); break;
				case 3: GetColumns(e.Node, context); break;
			}
		}

		private void GetDatabases(TreeNode node, Context context)
		{
			DataTree.UseWaitCursor = true;
			DataTree.Update();
			node.Nodes.Clear();
			if (!context.Has("Databases"))
			{
				TreeNode dbnode = new TreeNode(context.Database);
				dbnode.Tag = context.Database;
				node.Nodes.Add(dbnode);
				dbnode.Nodes.Add(new TreeNode());
			}
			else
				foreach (DataRow db in context.GetSchema("Databases").Rows)
				{
					try { context.Database = (string)db[0]; }
					catch (DbException dbex) { Trace.TraceError(dbex.ToString()); continue; }
					TreeNode dbnode = new TreeNode(context.Database);
					dbnode.Tag = context.Database;
					node.Nodes.Add(dbnode);
					dbnode.Nodes.Add(new TreeNode());
				}
			DataTree.UseWaitCursor = false;
		}

		private void GetTables(TreeNode node, Context context)
		{
			DataTree.UseWaitCursor = true;
			DataTree.Update();
			node.Nodes.Clear();
			foreach (string tabletype in new string[] { "Tables", "Views" })
			{
				if (!context.Has(tabletype)) continue;
				DataTable tables = context.Connection.GetSchema(tabletype);
				bool hasSchema = tables.Columns.Contains("TABLE_SCHEMA");
				DataRow[] usertables = (((tabletype == "Tables") && tables.Columns.Contains("TABLE_TYPE")) ?
					tables.Select("TABLE_TYPE in ('TABLE','BASE TABLE')") : tables.Select());
				foreach (DataRow table in usertables)
				{
					string catalog = (tables.Columns.Contains("TABLE_CATALOG") ?
						(string)table["TABLE_CATALOG"] : null);
					string schema = (hasSchema ? (string)table["TABLE_SCHEMA"] : null);
					string tablename = (string)table["TABLE_NAME"];
					TreeNode tablenode = new TreeNode(((schema == null) || (schema == "dbo")) ?
						tablename : schema + "." + tablename);
					tablenode.Tag = new string[] { catalog, schema, tablename };
					if ((tabletype == "Views") || ((string)table["TABLE_TYPE"] == "VIEW"))
					{
						tablenode.NodeFont = new Font(DataTree.Font, FontStyle.Italic);
						tablenode.ForeColor = Color.DarkBlue;
					}
					node.Nodes.Add(tablenode);
					tablenode.Nodes.Add(new TreeNode());
				}
			}
			DataTree.UseWaitCursor = false;
		}

		private void GetColumns(TreeNode node, Context context)
		{
			DataTree.UseWaitCursor = true;
			DataTree.Update();
			node.Nodes.Clear();
			foreach (DataRow row in context.Connection.GetSchema("Columns", (string[])node.Tag).Rows)
			{
				TreeNode colnode = new TreeNode((string)row["COLUMN_NAME"]);
				colnode.Tag = row;
				node.Nodes.Add(colnode);
			}
			DataTree.UseWaitCursor = false;
		}

		private void DataTree_AfterSelect(object sender, TreeViewEventArgs e)
		{
			AboutUI.Visible = ProviderUI.Visible = ConnUI.Visible = DbUI.Visible = TableUI.Visible = ColumnUI.Visible = false;
			TreeNode node = DataTree.SelectedNode;
			if (node.Tag == null) { AboutUI.Visible = true; AboutUI.BringToFront(); return; }
			Context context = null;
			try { context = new Context(node); }
			catch (DbException dbex)
			{
				Trace.TraceError(dbex.ToString());
				MessageBox.Show(dbex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				node.Remove();
				AboutUI.Visible = true;
				AboutUI.BringToFront();
				return;
			}
			if (node.Tag is ConnectionStringSettings)
			{
				try { node.Tag = new DistinguishedConnection(context.Factory, (ConnectionStringSettings)node.Tag); }
				catch (DbException dbex)
				{
					Trace.TraceError(dbex.ToString());
					MessageBox.Show(dbex.Message, "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
					node.Remove();
					AboutUI.Visible = true;
					AboutUI.BringToFront();
					return;
				}
			}
			switch (context.Depth)
			{
				case 0: ShowProviderPanel(context); break;
				case 1: ShowConnectionPanel(context); break;
				case 2: ShowDatabasePanel(context); break;
				case 3: ShowTablePanel(context); break;
				case 4: ShowColumnPanel(context); break;
			}
		}

		private void ShowProviderPanel(Context context)
		{
			DbConnectionStringBuilder connstr = context.Factory.CreateConnectionStringBuilder();
			ConnectionStringBox.Text = "";
			ConnectionStringValues.Rows.Clear();
			foreach (string key in connstr.Keys) ConnectionStringValues.Rows.Add(key, connstr[key]);
			ProviderUI.Visible = true;
			ProviderUI.BringToFront();
		}

		private void SyncConnectionStringBox(object sender, EventArgs e)
		{
			DbConnectionStringBuilder connstr = ((DbProviderFactory)DataTree.SelectedNode.Tag).CreateConnectionStringBuilder();
			foreach (DataGridViewRow row in ConnectionStringValues.Rows)
				if ((row.Cells[0].Value != null) && (row.Cells[0].Value.ToString() != "")
					&& (row.Cells[1].Value != null) && (row.Cells[1].Value.ToString() != ""))
					connstr[row.Cells[0].Value.ToString()] = row.Cells[1].Value.ToString();
			ConnectionStringBox.Text = connstr.ConnectionString;
		}

		private void SyncConnectionStringValues(object sender, EventArgs e)
		{
			DbConnectionStringBuilder connstr = ((DbProviderFactory)DataTree.SelectedNode.Tag).CreateConnectionStringBuilder();
			ConnectionStringValues.Rows.Clear();
			try { connstr.ConnectionString = ConnectionStringBox.Text; }
			catch (ArgumentException ex) { Trace.TraceWarning(ex.ToString()); }
			foreach (string key in connstr.Keys)
				ConnectionStringValues.Rows.Add(key, connstr[key]);
		}

		private void AddConnectionButton_Click(object sender, EventArgs e)
		{
			string name = Microsoft.VisualBasic.Interaction.InputBox("Choose a name for this connection:",
				"Connection Name", "", -1, -1),
				typename = DataTree.SelectedNode.Tag.GetType().FullName;
			if (String.IsNullOrEmpty(name)) return;
			RegistryKey dsns = Registry.CurrentUser.CreateSubKey(@"Software\webcoder\DBScope\" + typename);
			dsns.SetValue(name, ConnectionStringBox.Text);
			TreeNode newconn = new TreeNode(name);
			newconn.Tag = new ConnectionStringSettings(name, ConnectionStringBox.Text, typename);
			newconn.Nodes.Add(new TreeNode());
			DataTree.SelectedNode.Nodes.Add(newconn);
			DbConnectionStringBuilder connstr = ((DbProviderFactory)DataTree.SelectedNode.Tag).CreateConnectionStringBuilder();
			ConnectionStringBox.Text = "";
			ConnectionStringValues.Rows.Clear();
			foreach (string key in connstr.Keys)
				ConnectionStringValues.Rows.Add(key, connstr[key]);
		}

		private void ShowConnectionPanel(Context context)
		{
			if ((context.Info != null) && (context.Info.Rows.Count > 0))
			{
				DataTable infotable = FlipRow(context.Info.Rows[0]);
				infotable.Columns.RemoveAt(2);
				ConnectionInfo.DataSource = infotable;
				ConnectionInfo.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
				ConnectionInfo.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
				ConnectionInfo.Columns[1].FillWeight = 120;
				ConnectionInfo.Columns[1].DefaultCellStyle.Font = new Font("Lucida Console", 8.75F);
			}
			else ConnectionInfo.DataSource = null;
			ReservedWords.DataSource = context.ReservedWords;
			if ((context.ReservedWords != null) && (context.ReservedWords.Rows.Count > 0))
			{
				ReservedWords.Columns[0].HeaderText = "Reserved Words";
				ReservedWords.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
			}
			DataTypes.DataSource = context.Types;
			ConnUI.Visible = true;
			ConnUI.BringToFront();
		}

		private void ShowDatabasePanel(Context context)
		{
			QueryBox.Text = null;
			ResultsGrid.DataSource = null;
			ResultsTabs.SelectedIndex = 0;
			DbUI.Visible = true;
			DbUI.BringToFront();
		}

		private void ExecuteQueryButton_Click(object sender, EventArgs e)
		{
			ExecuteQuery();
		}

		private void ExecuteQuery()
		{
			DbUI.UseWaitCursor = true;
			DbUI.Update();
			TreeNode node = DataTree.SelectedNode;
			Context context = new Context(node);
			string query = QueryBox.Text;
			DataTable results = Execute(context.Connection, query);
			if ((context.Database != (string)node.Tag) &&
				(node.Parent.Nodes.Find(context.Database, false).Length == 0))
			{
				TreeNode newdbnode = new TreeNode(context.Database);
				newdbnode.Tag = context.Database;
				node.Parent.Nodes.Add(newdbnode);
				newdbnode.Nodes.Add(new TreeNode());
			}
			ResultsGrid.DataSource = results;
			QueryRowCount.Text = String.Format("{0:0;-0;No} rows", results.Rows.Count);
			ResultsTabs.SelectedIndex = 0;
			ResultsRowGrid.DataSource = null;
			DbUI.UseWaitCursor = false;
		}

		private void ResultsTabs_Selected(object sender, TabControlEventArgs e)
		{
			if (e.TabPageIndex != 1) return;
			if (ResultsGrid.CurrentRow == null) return;
			DataRowView row = ResultsGrid.CurrentRow.DataBoundItem as DataRowView;
			if (row == null) return;
			ResultsRowGrid.DataSource = FlipRow(row.Row);
			ResultsRowGrid.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
			ResultsRowGrid.Columns[0].Width = 150;
			ResultsRowGrid.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
			ResultsRowGrid.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
			ResultsRowGrid.Columns[2].Width = 150;
		}

		private void ResultsGrid_DoubleClick(object sender, EventArgs e)
		{
			ResultsTabs.SelectedIndex = 1;
		}

		private void ShowTablePanel(Context context)
		{
			TableTabs.SelectedIndex = 0;
			TableColumnsGrid.DataSource = context.GetColumns();
			TableSampleDataGrid.DataSource = null;
			TableDataGrid.DataSource = null;
			TableRowGrid.DataSource = null;
			DbCommand cmd = context.CreateCommand();
			cmd.CommandText = String.Format(ConfigurationManager.AppSettings["RowCountQuery"], context.TableName);
			try { TableRowCount.Text = String.Format("{0:0;-0;No} rows", cmd.ExecuteScalar()); }
			catch (DbException dbex) { TableRowCount.Text = ""; Trace.TraceWarning(dbex.ToString()); }
			TableUI.Visible = true;
			TableUI.BringToFront();
		}

		private void TableTabs_Selected(object sender, TabControlEventArgs e)
		{
			Context context = new Context(DataTree.SelectedNode);
			switch (TableTabs.SelectedIndex)
			{
				case 1:
					TableDataGrid.DataSource = null;
					TableSampleDataGrid.DataSource = Execute(context.Connection,
						String.Format(ConfigurationManager.AppSettings["TableSampleQuery"], context.TableName));
					break;
				case 2:
					TableSampleDataGrid.DataSource = null;
					TableDataGrid.DataSource = Execute(context.Connection,
						String.Format(ConfigurationManager.AppSettings["TableDataQuery"], context.TableName));
					break;
				case 3:
					DataGridViewRow gridrow = TableDataGrid.CurrentRow ?? TableSampleDataGrid.CurrentRow;
					if (gridrow == null) return;
					DataRowView row = gridrow.DataBoundItem as DataRowView;
					if (row == null) return;
					TableRowGrid.DataSource = FlipRow(row.Row);
					TableRowGrid.Columns[0].Width = 150;
					TableRowGrid.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
					TableRowGrid.Columns[2].Width = 150;
					break;
			}
		}

		private void TableSampleDataGrid_DoubleClick(object sender, EventArgs e)
		{
			TableTabs.SelectedIndex = 3;
		}

		private void TableDataGrid_DoubleClick(object sender, EventArgs e)
		{
			TableTabs.SelectedIndex = 3;
		}

		private void ShowColumnPanel(Context context)
		{
			if (DataTree.SelectedNode.Tag is DataRow)
			{
				DataTable stats = FlipRow(DataTree.SelectedNode.Tag as DataRow);
				foreach (string query in StatsQuery)
				{
					string q = String.Format(query, context.TableName, DataTree.SelectedNode.Text);
					try { stats.Merge(FlipRow(ExecuteNoCatch(context.Connection, q).Rows[0])); }
					catch (IndexOutOfRangeException) { }
					catch (DbException dbex) { Trace.TraceWarning("Query error: {0}\n{1}", q, dbex); }
				}
				stats.Columns.RemoveAt(2);
				DataTree.SelectedNode.Tag = stats;
			}
			ColumnStatsGrid.DataSource = DataTree.SelectedNode.Tag;
			ColumnStatsGrid.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
			ColumnStatsGrid.Columns[0].Width = 150;
			ColumnStatsGrid.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
			ColumnTabs.SelectedIndex = 0;
			ColumnDataGrid.DataSource = null;
			ColumnUI.Visible = true;
			ColumnUI.BringToFront();
		}

		private void ColumnTabs_Selected(object sender, TabControlEventArgs e)
		{
			Context context = new Context(DataTree.SelectedNode);
			if (ColumnTabs.SelectedIndex == 1)
			{
				try
				{
					ColumnDataGrid.DataSource = ExecuteNoCatch(context.Connection,
						String.Format(ConfigurationManager.AppSettings["ColumnDataQuery"],
						context.TableName, DataTree.SelectedNode.Text));
				}
				catch (DbException)
				{ // simplify query for datatypes that do not support comparison, such as text, image, memo
					ColumnDataGrid.DataSource = Execute(context.Connection,
						String.Format(ConfigurationManager.AppSettings["SimpleColumnDataQuery"],
						context.TableName, DataTree.SelectedNode.Text));
				}
				ColumnDataGrid.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
				ColumnDataGrid.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
				ColumnDataGrid.Columns[1].Width = 60;
			}
		}

		private void SiteLink_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
		{
			try
			{
				Process.Start(SiteLink.Text);
				SiteLink.LinkVisited = true;
			}
			catch (Win32Exception ex)
			{
				MessageBox.Show(ex.Message, "Link Prevented", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void LicenseLink_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
		{
			try
			{
				Process.Start((string)LicenseLink.Tag);
				LicenseLink.LinkVisited = true;
			}
			catch (Win32Exception ex)
			{
				MessageBox.Show(ex.Message, "Link Prevented", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private DataTable Execute(DbConnection conn, string query)
		{
			DataTable data = new DataTable();
			DbCommand cmd = conn.CreateCommand();
			cmd.CommandText = query;
			try { data.Load(cmd.ExecuteReader()); }
			catch (DbException e) { MessageBox.Show(e.Message, "Query Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
			return data;
		}

		private DataTable ExecuteNoCatch(DbConnection conn, string query)
		{
			DataTable data = new DataTable();
			DbCommand cmd = conn.CreateCommand();
			cmd.CommandText = query;
			data.Load(cmd.ExecuteReader());
			return data;
		}

		private DataTable FlipRow(DataRow row)
		{
			DataTable table = new DataTable();
			table.Columns.Add("Name");
			table.Columns.Add("Value");
			table.Columns.Add("Type");
			foreach (DataColumn col in row.Table.Columns)
				table.Rows.Add(new object[] { col.ColumnName, row[col], col.DataType });
			return table;
		}
	}
}