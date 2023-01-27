#define OLE_PROPERTY

using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OpenMcdf;
using OpenMcdf.Extensions;
using StructuredStorageExplorer.Properties;

// Author Federico Blaseotto

namespace StructuredStorageExplorer
{

    /// <summary>
    /// Sample Structured Storage viewer to 
    /// demonstrate use of OpenMCDF
    /// </summary>
    public partial class MainForm : Form
    {
        private CompoundFile _cf;
        private FileStream _fs;

        public MainForm()
        {
            InitializeComponent();

#if !OLE_PROPERTY
            tabControl1.TabPages.Remove(tabPage2);
#endif

            //Load images for icons from resx
            var folderImage = (Image)Resources.ResourceManager.GetObject("storage");
            var streamImage = (Image)Resources.ResourceManager.GetObject("stream");
            //Image olePropsImage = (Image)Properties.Resources.ResourceManager.GetObject("oleprops");

            treeView1.ImageList = new ImageList();
            treeView1.ImageList.Images.Add(folderImage);
            treeView1.ImageList.Images.Add(streamImage);
            //treeView1.ImageList.Images.Add(olePropsImage);



            saveAsToolStripMenuItem.Enabled = false;
            updateCurrentFileToolStripMenuItem.Enabled = false;

        }



        private void OpenFile()
        {
            if (!string.IsNullOrEmpty(openFileDialog1.FileName))
            {
                CloseCurrentFile();

                treeView1.Nodes.Clear();
                fileNameLabel.Text = openFileDialog1.FileName;
                LoadFile(openFileDialog1.FileName, true);
                _canUpdate = true;
                saveAsToolStripMenuItem.Enabled = true;
                updateCurrentFileToolStripMenuItem.Enabled = true;
            }
        }

        private void CloseCurrentFile()
        {
            if (_cf != null)
                _cf.Close();

            if (_fs != null)
                _fs.Close();

            treeView1.Nodes.Clear();
            fileNameLabel.Text = string.Empty;
            saveAsToolStripMenuItem.Enabled = false;
            updateCurrentFileToolStripMenuItem.Enabled = false;

            propertyGrid1.SelectedObject = null;
            hexEditor.ByteProvider = null;
        }

        private bool _canUpdate;

        private void CreateNewFile()
        {
            CloseCurrentFile();

            _cf = new CompoundFile();
            _canUpdate = false;
            saveAsToolStripMenuItem.Enabled = true;

            updateCurrentFileToolStripMenuItem.Enabled = false;

            RefreshTree();
        }

        private void RefreshTree()
        {
            treeView1.Nodes.Clear();

            TreeNode root = null;
            root = treeView1.Nodes.Add("Root Entry", "Root");
            root.ImageIndex = 0;
            root.Tag = _cf.RootStorage;

            //Recursive function to get all storage and streams
            AddNodes(root, _cf.RootStorage);
        }

        private void LoadFile(string fileName, bool enableCommit)
        {

            _fs = new FileStream(
                fileName,
                FileMode.Open,
                enableCommit ?
                    FileAccess.ReadWrite
                    : FileAccess.Read
                );

            try
            {
                if (_cf != null)
                {
                    _cf.Close();
                    _cf = null;
                }

                //Load file
                if (enableCommit)
                {
                    _cf = new CompoundFile(_fs, CfsUpdateMode.Update, CfsConfiguration.SectorRecycle | CfsConfiguration.NoValidationException | CfsConfiguration.EraseFreeSectors);
                }
                else
                {
                    _cf = new CompoundFile(_fs);
                }

                RefreshTree();
            }
            catch (Exception ex)
            {
                treeView1.Nodes.Clear();
                fileNameLabel.Text = string.Empty;
                MessageBox.Show("Internal error: " + ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Recursive addition of tree nodes foreach child of current item in the storage
        /// </summary>
        /// <param name="node">Current TreeNode</param>
        /// <param name="cfs">Current storage associated with node</param>
        private void AddNodes(TreeNode node, CfStorage cfs)
        {
            var va = delegate (CfItem target)
            {
                var temp = node.Nodes.Add(
                    target.Name,
                    target.Name + (target.IsStream ? " (" + target.Size + " bytes )" : "")
                    );

                temp.Tag = target;

                if (target.IsStream)
                {

                    //Stream
                    temp.ImageIndex = 1;
                    temp.SelectedImageIndex = 1;
                }
                else
                {
                    //Storage
                    temp.ImageIndex = 0;
                    temp.SelectedImageIndex = 0;

                    //Recursion into the storage
                    AddNodes(temp, (CfStorage)target);
                }
            };

            //Visit NON-recursively (first level only)
            cfs.VisitEntries(va, false);
        }



        private void exportDataToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //No export if storage
            if (treeView1.SelectedNode == null || !((CfItem)treeView1.SelectedNode.Tag).IsStream)
            {
                MessageBox.Show("Only stream data can be exported", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                return;
            }

            var target = (CfStream)treeView1.SelectedNode.Tag;

            // A lot of stream and storage have only non-printable characters.
            // We need to sanitize filename.

            var sanitizedFileName = string.Empty;

            foreach (var c in target.Name)
            {
                if (
                    char.GetUnicodeCategory(c) == UnicodeCategory.LetterNumber
                    || char.GetUnicodeCategory(c) == UnicodeCategory.LowercaseLetter
                    || char.GetUnicodeCategory(c) == UnicodeCategory.UppercaseLetter
                    )

                    sanitizedFileName += c;
            }

            if (string.IsNullOrEmpty(sanitizedFileName))
            {
                sanitizedFileName = "tempFileName";
            }

            saveFileDialog1.FileName = sanitizedFileName + ".bin";

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileStream fs = null;

                try
                {
                    fs = new FileStream(saveFileDialog1.FileName, FileMode.CreateNew, FileAccess.ReadWrite);
                    fs.Write(target.GetData(), 0, (int)target.Size);
                }
                catch (Exception ex)
                {
                    treeView1.Nodes.Clear();
                    MessageBox.Show("Internal error: " + ex.Message, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    if (fs != null)
                    {
                        fs.Flush();
                        fs.Close();
                        fs = null;
                    }
                }
            }
        }

        private void removeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var n = treeView1.SelectedNode;
            ((CfStorage)n.Parent.Tag).Delete(n.Name);

            RefreshTree();
        }

        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog1.FilterIndex = 2;
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                _cf.Save(saveFileDialog1.FileName);
            }
        }

        private void updateCurrentFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (_canUpdate)
            {
                if (hexEditor.ByteProvider != null && hexEditor.ByteProvider.HasChanges())
                    hexEditor.ByteProvider.ApplyChanges();
                _cf.Commit();
            }
            else
                MessageBox.Show("Cannot update a compound document that is not based on a stream or on a file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

        }

        private void addStreamToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var streamName = string.Empty;

            if (Utils.InputBox("Add stream", "Insert stream name", ref streamName) == DialogResult.OK)
            {
                var cfs = treeView1.SelectedNode.Tag as CfItem;

                if (cfs != null && (cfs.IsStorage || cfs.IsRoot))
                {
                    try
                    {
                        ((CfStorage)cfs).AddStream(streamName);
                    }
                    catch (CfDuplicatedItemException)
                    {
                        MessageBox.Show("Cannot insert a duplicated item", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }


                RefreshTree();
            }
        }

        private void addStorageStripMenuItem1_Click(object sender, EventArgs e)
        {
            var storage = string.Empty;

            if (Utils.InputBox("Add storage", "Insert storage name", ref storage) == DialogResult.OK)
            {
                var cfs = treeView1.SelectedNode.Tag as CfItem;

                if (cfs != null && (cfs.IsStorage || cfs.IsRoot))
                {
                    try
                    {
                        ((CfStorage)cfs).AddStorage(storage);
                    }
                    catch (CfDuplicatedItemException)
                    {
                        MessageBox.Show("Cannot insert a duplicated item", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

                RefreshTree();
            }
        }

        private void importDataStripMenuItem1_Click(object sender, EventArgs e)
        {
            var fileName = string.Empty;

            if (openDataFileDialog.ShowDialog() == DialogResult.OK)
            {
                var s = treeView1.SelectedNode.Tag as CfStream;

                if (s != null)
                {
                    var f = new FileStream(openDataFileDialog.FileName, FileMode.Open, FileAccess.Read, FileShare.Read);
                    var data = new byte[f.Length];
                    f.Read(data, 0, (int)f.Length);
                    f.Flush();
                    f.Close();
                    s.SetData(data);

                    RefreshTree();
                }
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_cf != null)
                _cf.Close();
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

        }

        private void newStripMenuItem1_Click(object sender, EventArgs e)
        {

            CreateNewFile();
        }

        private void openFileMenuItem_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    OpenFile();
                }
                catch
                {

                }
            }
        }


        private void treeView1_MouseUp(object sender, MouseEventArgs e)
        {
            // Get the node under the mouse cursor.
            // We intercept both left and right mouse clicks
            // and set the selected treenode according.

            var n = treeView1.GetNodeAt(e.X, e.Y);

            if (n != null)
            {
                if (hexEditor.ByteProvider != null && hexEditor.ByteProvider.HasChanges())
                {
                    if (MessageBox.Show("Do you want to save pending changes ?", "Save changes", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                    {
                        hexEditor.ByteProvider.ApplyChanges();
                    }
                }

                treeView1.SelectedNode = n;


                // The tag property contains the underlying CFItem.
                var target = (CfItem)n.Tag;

                if (target.IsStream)
                {
                    addStorageStripMenuItem1.Enabled = false;
                    addStreamToolStripMenuItem.Enabled = false;
                    importDataStripMenuItem1.Enabled = true;
                    exportDataToolStripMenuItem.Enabled = true;

#if OLE_PROPERTY
                    if (target.Name == "\u0005SummaryInformation" || target.Name == "\u0005DocumentSummaryInformation")
                    {
                        var c = ((CfStream)target).AsOlePropertiesContainer();

                        var ds = new DataTable();

                        ds.Columns.Add("Name", typeof(string));
                        ds.Columns.Add("Type", typeof(string));
                        ds.Columns.Add("Value", typeof(string));

                        foreach (var p in c.Properties)
                        {
                            if (p.Value.GetType() != typeof(byte[]) && p.Value.GetType().GetInterfaces().Any(t => t == typeof(IList)))
                            {
                                for (var h = 0; h < ((IList)p.Value).Count; h++)
                                {
                                    var dr = ds.NewRow();
                                    dr.ItemArray = new[] { p.PropertyName, p.VtType, ((IList)p.Value)[h] };
                                    ds.Rows.Add(dr);
                                }
                            }
                            else
                            {
                                var dr = ds.NewRow();
                                dr.ItemArray = new[] { p.PropertyName, p.VtType, p.Value };
                                ds.Rows.Add(dr);
                            }
                        }
                        ds.AcceptChanges();
                        dgvOLEProps.DataSource = ds;

                        if (c.HasUserDefinedProperties)
                        {
                            var ds2 = new DataTable();

                            ds2.Columns.Add("Name", typeof(string));
                            ds2.Columns.Add("Type", typeof(string));
                            ds2.Columns.Add("Value", typeof(string));

                            foreach (var p in c.UserDefinedProperties.Properties)
                            {
                                if (p.Value.GetType() != typeof(byte[]) && p.Value.GetType().GetInterfaces().Any(t => t == typeof(IList)))
                                {
                                    for (var h = 0; h < ((IList)p.Value).Count; h++)
                                    {
                                        var dr = ds2.NewRow();
                                        dr.ItemArray = new[] { p.PropertyName, p.VtType, ((IList)p.Value)[h] };
                                        ds2.Rows.Add(dr);
                                    }
                                }
                                else
                                {
                                    var dr = ds2.NewRow();
                                    dr.ItemArray = new[] { p.PropertyName, p.VtType, p.Value };
                                    ds2.Rows.Add(dr);
                                }
                            }

                            ds2.AcceptChanges();
                            dgvUserDefinedProperties.DataSource = ds2;
                        }

                       
                    }
                    else
                    {
                        dgvOLEProps.DataSource = null;
                    }

                    //if (target.Name == "\u0005SummaryInformation" || target.Name == "\u0005DocumentSummaryInformation")
                    //{
                    //    ContainerType map = target.Name == "\u0005SummaryInformation" ? ContainerType.SummaryInfo : ContainerType.DocumentSummaryInfo;
                    //    PropertySetStream mgr = ((CFStream)target).AsOLEProperties();

                    //    DataTable ds = new DataTable();
                    //    ds.Columns.Add("Name", typeof(String));
                    //    ds.Columns.Add("Type", typeof(String));
                    //    ds.Columns.Add("Value", typeof(String));

                    //    for (int i = 0; i < mgr.PropertySet0.NumProperties; i++)
                    //    {
                    //        ITypedPropertyValue p = mgr.PropertySet0.Properties[i];

                    //        if (p.Value.GetType().GetInterfaces().Any(t => t == typeof(IList)))
                    //        {
                    //            for (int h = 0; h < ((IList)p.Value).Count; h++)
                    //            {
                    //                DataRow dr = ds.NewRow();
                    //                dr.ItemArray = new Object[] { mgr.PropertySet0.PropertyIdentifierAndOffsets[i].PropertyIdentifier.GetDescription(map), p.VTType, ((IList)p.Value)[h] };
                    //                ds.Rows.Add(dr);
                    //            }
                    //        }
                    //        else
                    //        {
                    //            DataRow dr = ds.NewRow();
                    //            dr.ItemArray = new Object[] { mgr.PropertySet0.PropertyIdentifierAndOffsets[i].PropertyIdentifier.GetDescription(map), p.VTType, p.Value };
                    //            ds.Rows.Add(dr);
                    //        }
                    //    }

                    //    ds.AcceptChanges();
                    //    dgvOLEProps.DataSource = ds;
                    //}
#endif
                }
            }
            else
            {
                addStorageStripMenuItem1.Enabled = true;
                addStreamToolStripMenuItem.Enabled = true;
                importDataStripMenuItem1.Enabled = false;
                exportDataToolStripMenuItem.Enabled = false;
            }

            if (n != null)
                propertyGrid1.SelectedObject = n.Tag;


            if (n != null)
            {
                var targetStream = n.Tag as CfStream;
                if (targetStream != null)
                {
                    hexEditor.ByteProvider = new StreamDataProvider(targetStream);
                }
                else
                {
                    hexEditor.ByteProvider = null;
                }
            }
        }

        void hexEditor_ByteProviderChanged(object sender, EventArgs e)
        {

        }

        private void closeStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (hexEditor.ByteProvider != null && hexEditor.ByteProvider.HasChanges())
            {
                if (MessageBox.Show("Do you want to save pending changes ?", "Save changes", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    hexEditor.ByteProvider.ApplyChanges();
                }
            }

            CloseCurrentFile();
        }


    }
}
