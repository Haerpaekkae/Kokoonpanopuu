using System;
using System.Collections;
using System.Drawing;
using System.Windows.Forms;
using Tekla.Technology.Akit.UserScript;
using Tekla.Technology.Scripting;
using Tekla.Structures;
using TSG = Tekla.Structures.Geometry3d;
using Tekla.Structures.Model;
using Tekla.Structures.Model.Operations;
using UI = Tekla.Structures.Model.UI;

namespace Tekla.Technology.Akit.UserScript
{
    public partial class AssemblyTree : Form
    {
        private TreeClipBoard _clipBoard = new TreeClipBoard();
        private Model _tsModel = new Model();
        private Events events = new Events();
        private TreeView _activeTree = null;
        private TreeNode _hoveredNode = null;
        private Assembly _rootAssembly = null;
        private static bool preventModelSelectionEvent = false;
        delegate void GetSelectedCallback();
        delegate void CloseCallback();

        public AssemblyTree()
        {
            InitializeComponent();
            //InitializeForm();

            GetSelected();

            events.Register();
            events.SelectionChange += new Tekla.Structures.Model.Events.SelectionChangeDelegate(events_SelectionChange);
            events.TeklaStructuresExit += new Events.TeklaStructuresExitDelegate(events_TS_Exit);
        }

        void events_TS_Exit()
        {
            if (treeView1.InvokeRequired)
                Invoke(new CloseCallback(this.Close));
            else
                this.Close();
        }

        void events_SelectionChange()
        {
            if (!preventModelSelectionEvent)
            {
                if (treeView1.InvokeRequired)
                    Invoke(new GetSelectedCallback(GetSelected));
                else
                    GetSelected();
            }
        }

        private void GetSelected()
        {
            if (!checkBoxLockTrees.Checked)
            {
                treeView1.Nodes.Clear();
                treeView2.Nodes.Clear();
                HideTree2();
                toolStripButtonCure.Visible = false;

                int index = 0;
                UI.ModelObjectSelector selector = new UI.ModelObjectSelector();
                ModelObjectEnumerator selectedEnum = selector.GetSelectedObjects();

                while (index < 2 && selectedEnum.MoveNext())
                {
                    Assembly assembly = null;

                    try
                    {
                        assembly = selectedEnum.Current as Assembly;
                    }
                    catch
                    {
                    }

                    if (assembly != null)
                    {
                        if (index == 0)
                        {
                            FillTree(assembly, treeView1);
                        }
                        else
                        {
                            FillTree(assembly, treeView2);
                            ShowTree2();
                            toolStripButtonCure.Visible = false;
                        }
                    }

                    index++;
                }
            }
            else
            {
                UI.ModelObjectSelector selector = new UI.ModelObjectSelector();
                ModelObjectEnumerator selectedEnum = selector.GetSelectedObjects();

                if(selectedEnum.MoveNext())
                {
                    Identifier identifier = (selectedEnum.Current as ModelObject).Identifier;

                    TreeNode[] nodes1 = treeView1.Nodes.Find(identifier.ID.ToString(), true);
                    TreeNode[] nodes2 = treeView2.Nodes.Find(identifier.ID.ToString(), true);

                    if (nodes1 != null && nodes1.Length > 0)
                    {
                        treeView1.SelectedNode = nodes1[0];
                        treeView1.Focus();
                    }
                    else if (nodes2 != null && nodes2.Length > 0)
                    {
                        treeView2.SelectedNode = nodes2[0];
                        treeView2.Focus();
                    }
                }
            }
        }

        private bool CheckValidity(Part part)
        {
            bool result = true;
            string partMaterial = "";
            string rootMaterial = "";
            
            part.GetReportProperty("MATERIAL_TYPE", ref partMaterial);
            _rootAssembly.GetMainPart().GetReportProperty("MATERIAL_TYPE", ref rootMaterial);

            if (rootMaterial == "CONCRETE" &&
                partMaterial != "CONCRETE" &&
                part.GetAssembly().Identifier.ID == _rootAssembly.Identifier.ID)
            {
                result = false;
            }

            return result;
        }

        private void AddReinforcements(Part part, TreeNode partNode)
        {
            ModelObjectEnumerator rebarEnum = part.GetReinforcements();

            while (rebarEnum.MoveNext())
            {
                Reinforcement rebar = null;

                try
                {
                    rebar = rebarEnum.Current as Reinforcement;
                }
                catch
                {
                }

                if (rebar != null)
                {
                    TreeNode childNode = new TreeNode(rebar.Name);
                    childNode.Tag = rebar.Identifier;
                    childNode.Name = rebar.Identifier.ID.ToString();
                    partNode.Nodes.Add(childNode);
                }
            }
        }

        private void AddAssemblyLeafs(Assembly assembly, TreeNode assemblyNode)
        {
            string assPos = "";

            Part mainPart = assembly.GetMainPart() as Part;

            if (mainPart == null)
            {
                TreeNode mainPartNode = new TreeNode("EI PÄÄOSAA!" + "   " + assPos);
                mainPartNode.BackColor = Color.Orange;
                mainPartNode.Expand();
                assemblyNode.Nodes.Add(mainPartNode);
            }
            else
            {
                mainPart.GetReportProperty("ASSEMBLY.ASSEMBLY_POS", ref assPos);
                //mainPart.GetReportProperty("CAST_UNIT.CAST_UNIT_POS", ref assPos);

                TreeNode mainPartNode = new TreeNode(mainPart.Name + "   " + assPos);
                mainPartNode.Tag = mainPart.Identifier;
                mainPartNode.Name = mainPart.Identifier.ID.ToString();
                assemblyNode.Nodes.Add(mainPartNode);

                ArrayList secondaries = assembly.GetSecondaries();

                foreach (Part part in secondaries)
                {
                    if (part != null)
                    {
                        assPos = "";
                        part.GetReportProperty("PART_POS", ref assPos);
                        //part.GetReportProperty("CAST_UNIT.CAST_UNIT_POS", ref assPos);

                        TreeNode partNode = new TreeNode(part.Name + "   " + assPos);
                        partNode.Tag = part.Identifier;
                        partNode.Name = part.Identifier.ID.ToString();

                        if (!CheckValidity(part))
                        {
                            partNode.BackColor = Color.Red;
                            toolStripButtonCure.Visible = true;
                        }

                        AddReinforcements(part, partNode);
                        assemblyNode.Nodes.Add(partNode);
                    }
                }

                AddReinforcements(mainPart, mainPartNode);
            }
        }

        private void FillTree(Assembly assembly, TreeView treeView)
        {
            _rootAssembly = assembly;

            string assPos = "";
            assembly.GetReportProperty("ASSEMBLY_POS", ref assPos);

            TreeNode node = new TreeNode(assembly.Name + "   " + assPos + "                   ", GetAssemblyChildren(assembly));
            node.Tag = assembly.Identifier;
            node.Name = assembly.Identifier.ID.ToString();
            
            node.NodeFont = new Font(FontFamily.GenericSansSerif, (float)10.0);
            node.NodeFont = new Font(node.NodeFont, FontStyle.Bold);

            AddAssemblyLeafs(assembly, node);

            treeView.Nodes.Add(node);
            treeView.Nodes[0].Expand();
            ExpandNumbNodes(treeView.Nodes);
        }

        private void ExpandNumbNodes(TreeNodeCollection nodes)
        {
            for (int ii = 0; ii < nodes.Count; ii++)
            {
                if (nodes[ii].Name == "")
                    nodes[ii].Parent.Expand();

                ExpandNumbNodes(nodes[ii].Nodes);
            }
        }

        TreeNode[] GetAssemblyChildren(Assembly assembly)
        {
            TreeNode[] result = null;
            ArrayList resultArray = new ArrayList();

            ArrayList subs = assembly.GetSubAssemblies();

            foreach (Assembly subAssembly in subs)
            {
                string assPos = "";
                subAssembly.GetReportProperty("ASSEMBLY_POS", ref assPos);

                TreeNode node = new TreeNode(subAssembly.Name + "   " + assPos, GetAssemblyChildren(subAssembly));
                node.Tag = subAssembly.Identifier;
                node.Name = subAssembly.Identifier.ID.ToString();

                node.NodeFont = new Font(FontFamily.GenericSansSerif, (float)10.0);
                node.NodeFont = new Font(node.NodeFont, FontStyle.Bold);

                AddAssemblyLeafs(subAssembly, node);

                resultArray.Add(node);
            }

            result = resultArray.ToArray(typeof(TreeNode)) as TreeNode[];
            return result;
        }

        private void RefreshTree(TreeView treeView)
        {
            int ID = Convert.ToInt32(treeView.Nodes[0].Name);
            Identifier rootAssemblyID = new Identifier(ID); //treeView.Nodes[0].Tag as Identifier;
            Assembly assembly = _tsModel.SelectModelObject(rootAssemblyID) as Assembly;

            toolStripButtonCure.Visible = false;

            if (assembly != null)
            {
                treeView.Nodes.Clear();
                FillTree(assembly, treeView);
            }
        }

        private void HideTree2()
        {
            treeView2.Visible = false;
            this.Size = new Size(300, this.Size.Height);
            treeView1.Size = new Size(246, treeView1.Size.Height);
        }

        private void ShowTree2()
        {
            treeView2.Visible = true;
            this.Size = new Size(555, this.Size.Height);
            treeView1.Size = new Size(246, treeView1.Size.Height);
        }

        private void treeView_AfterSelect(object sender, TreeViewEventArgs e)
        {
            Identifier id = e.Node.Tag as Identifier;
            Beam beam = new Beam();
            beam.Identifier = id;

            ArrayList mObjects = new ArrayList();
            mObjects.Add(beam);
            UI.ModelObjectSelector selector = new UI.ModelObjectSelector();

            preventModelSelectionEvent = true;
            selector.Select(mObjects);
            preventModelSelectionEvent = false;
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            events.UnRegister();
        }

        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                _activeTree = treeView1;
                contextMenuStrip1.Show(new Point(e.X + this.Location.X + this.treeView1.Bounds.X + 50, e.Y  + this.Location.Y+ this.treeView1.Bounds.Y));
            }
            else
            {
                Identifier id = new Identifier(Convert.ToInt32(e.Node.Name)); //e.Node.Tag as Identifier;
                Beam beam = new Beam();
                beam.Identifier = id;

                ArrayList mObjects = new ArrayList();
                mObjects.Add(beam);
                UI.ModelObjectSelector selector = new UI.ModelObjectSelector();

                preventModelSelectionEvent = true;
                selector.Select(mObjects);
                preventModelSelectionEvent = false;
            }
        }

        private void treeView2_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                _activeTree = treeView2;
                contextMenuStrip1.Show(new Point(e.X + this.Location.X + this.treeView2.Bounds.X + 50, e.Y + this.Location.Y + this.treeView2.Bounds.Y));
            }
            else
            {
                Identifier id = new Identifier(Convert.ToInt32(e.Node.Name)); //e.Node.Tag as Identifier;
                Beam beam = new Beam();
                beam.Identifier = id;

                ArrayList mObjects = new ArrayList();
                mObjects.Add(beam);
                UI.ModelObjectSelector selector = new UI.ModelObjectSelector();

                preventModelSelectionEvent = true;
                selector.Select(mObjects);
                preventModelSelectionEvent = false;
            }
        }

        private void contextMenuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            //_activeTree.SelectedNode.Tag as Identifier
            Identifier selectedID = new Identifier(Convert.ToInt32(_activeTree.SelectedNode.Name));

            switch (e.ClickedItem.Name)
            {
                case "copyToolStripMenuItem":
                    _clipBoard = new TreeClipBoard(TreeOperationEnum.COPY);
                    _clipBoard.AddObjectToClipBoard(selectedID);
                    break;
                case "copyChildrenToolStripMenuItem":
                    _clipBoard = new TreeClipBoard();
                    _clipBoard.AddChildrenToClipBoard(selectedID);
                    break;
                case "cutToolStripMenuItem":
                    _clipBoard = new TreeClipBoard(TreeOperationEnum.CUT);
                    _clipBoard.AddObjectToClipBoard(selectedID);
                    break;
                case "pasteToolStripMenuItem":
                    _clipBoard.DumpClipBoardHere(selectedID);

                    if(_clipBoard.CurrentOperation == TreeOperationEnum.CUT)
                        _clipBoard = new TreeClipBoard(); //Cut only once

                    RefreshTree(_activeTree);
                    break;
            }
        }

        private void toolStripButtonCure_Click(object sender, EventArgs e)
        {
            Assembly assembly = new Assembly();
            assembly.Identifier = new Identifier(Convert.ToInt32(treeView1.Nodes[0].Name));
            assembly.Select();

            ArrayList secondaries = assembly.GetSecondaries();

            for (int ii = 0; ii < secondaries.Count; ii++)
            {
                Part part = secondaries[ii] as Part;

                if (part != null)
                {
                    string materialType = "";
                    part.GetReportProperty("MATERIAL_TYPE", ref materialType);

                    if (materialType != "CONCRETE")
                    {
                        RemoveWeldsAndBolts(part);

                        assembly.Remove(part);
                        assembly.Modify();
                        ii--;

                        Assembly partAssembly = part.GetAssembly();
                        partAssembly.Name = part.Name;
                        partAssembly.Modify();

                        assembly.Add(partAssembly);
                        assembly.Modify();
                    }
                }
            }

            _tsModel.CommitChanges();
            RefreshTree(treeView1);
        }

        public static void RemoveWeldsAndBolts(Part part)
        {
            ModelObjectEnumerator weldEnum = part.GetWelds();
            ModelObjectEnumerator boltEnum = part.GetBolts();

            while (weldEnum.MoveNext())
                weldEnum.Current.Delete();

            while (boltEnum.MoveNext())
                boltEnum.Current.Delete();
        }

        private void treeView1_DragDrop(object sender, DragEventArgs e)
        {
            treeView1.SelectedNode = FindTreeNode(treeView1.Nodes, new Point(e.X, e.Y));
            Identifier selectedID = new Identifier(Convert.ToInt32(treeView1.SelectedNode.Name));

            _clipBoard.DumpClipBoardHere(selectedID);
            _clipBoard = new TreeClipBoard();
            RefreshTree(treeView1);
        }

        private void treeView1_ItemDrag(object sender, ItemDragEventArgs e)
        {
            Identifier selectedID = new Identifier(Convert.ToInt32(treeView1.SelectedNode.Name));
            _clipBoard = new TreeClipBoard(TreeOperationEnum.CUT);
            _clipBoard.AddObjectToClipBoard(selectedID);

            this.DoDragDrop("Siirrä...", DragDropEffects.Move);
        }

        private void treeView1_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void treeView1_DragOver(object sender, DragEventArgs e)
        {
            if (_hoveredNode != null)
            {
                _hoveredNode.BackColor = Color.White;
                _hoveredNode.ForeColor = Color.Black;
            }

            _hoveredNode = FindTreeNode(treeView1.Nodes, new Point(e.X, e.Y));
            _hoveredNode.BackColor = Color.Blue;
            _hoveredNode.ForeColor = Color.White;
        }

        private void treeView1_DragLeave(object sender, EventArgs e)
        {
            //_hoveredNode.BackColor = Color.White;
            //_hoveredNode.ForeColor = Color.Black;
        }

        private TreeNode FindTreeNode(TreeNodeCollection nodes, Point location)
        {
            TreeNode result = null;

            for (int ii = 0; result == null && ii < nodes.Count; ii++)
            {
                Rectangle bounds = new Rectangle(this.Bounds.X + treeView1.Bounds.X + nodes[ii].Bounds.X,
                                                 this.Bounds.Y + treeView1.Bounds.Y + nodes[ii].Bounds.Y + 31,
                                                 nodes[ii].Bounds.Width, nodes[ii].Bounds.Height);

                if (bounds.IntersectsWith(new Rectangle(location, new Size(1, 1))))
                    result = nodes[ii];

                if(result == null)
                    result = FindTreeNode(nodes[ii].Nodes, location);
            }

            return result;
        }
    }

    partial class AssemblyTree
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.treeView1 = new System.Windows.Forms.TreeView();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.cutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.copyToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.pasteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.treeView2 = new System.Windows.Forms.TreeView();
            this.checkBoxLockTrees = new System.Windows.Forms.CheckBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.toolStripButtonCure = new System.Windows.Forms.ToolStripButton();
            this.contextMenuStrip1.SuspendLayout();
            this.toolStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // treeView1
            // 
            this.treeView1.AllowDrop = true;
            this.treeView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.treeView1.Location = new System.Drawing.Point(16, 37);
            this.treeView1.Margin = new System.Windows.Forms.Padding(4);
            this.treeView1.Name = "treeView1";
            this.treeView1.Size = new System.Drawing.Size(246, 612);
            this.treeView1.TabIndex = 0;
            this.treeView1.ItemDrag += new System.Windows.Forms.ItemDragEventHandler(this.treeView1_ItemDrag);
            this.treeView1.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.treeView_AfterSelect);
            this.treeView1.NodeMouseClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.treeView1_NodeMouseClick);
            this.treeView1.DragDrop += new System.Windows.Forms.DragEventHandler(this.treeView1_DragDrop);
            this.treeView1.DragEnter += new System.Windows.Forms.DragEventHandler(this.treeView1_DragEnter);
            this.treeView1.DragOver += new System.Windows.Forms.DragEventHandler(this.treeView1_DragOver);
            this.treeView1.DragLeave += new System.EventHandler(this.treeView1_DragLeave);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.cutToolStripMenuItem,
            this.copyToolStripMenuItem,
            this.pasteToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(128, 76);
            this.contextMenuStrip1.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.contextMenuStrip1_ItemClicked);
            // 
            // cutToolStripMenuItem
            // 
            this.cutToolStripMenuItem.Name = "cutToolStripMenuItem";
            this.cutToolStripMenuItem.Size = new System.Drawing.Size(127, 24);
            this.cutToolStripMenuItem.Text = "Leikkaa";
            // 
            // copyToolStripMenuItem
            // 
            this.copyToolStripMenuItem.Name = "copyToolStripMenuItem";
            this.copyToolStripMenuItem.Size = new System.Drawing.Size(127, 24);
            this.copyToolStripMenuItem.Text = "Kopioi";
            // 
            // pasteToolStripMenuItem
            // 
            this.pasteToolStripMenuItem.Name = "pasteToolStripMenuItem";
            this.pasteToolStripMenuItem.Size = new System.Drawing.Size(127, 24);
            this.pasteToolStripMenuItem.Text = "Liitä";
            // 
            // treeView2
            // 
            this.treeView2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.treeView2.Location = new System.Drawing.Point(283, 37);
            this.treeView2.Margin = new System.Windows.Forms.Padding(4);
            this.treeView2.Name = "treeView2";
            this.treeView2.Size = new System.Drawing.Size(241, 612);
            this.treeView2.TabIndex = 1;
            this.treeView2.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.treeView_AfterSelect);
            this.treeView2.NodeMouseClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.treeView2_NodeMouseClick);
            // 
            // checkBoxLockTrees
            // 
            this.checkBoxLockTrees.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.checkBoxLockTrees.AutoSize = true;
            this.checkBoxLockTrees.Location = new System.Drawing.Point(16, 656);
            this.checkBoxLockTrees.Name = "checkBoxLockTrees";
            this.checkBoxLockTrees.Size = new System.Drawing.Size(107, 21);
            this.checkBoxLockTrees.TabIndex = 2;
            this.checkBoxLockTrees.Text = "Lukitse puut";
            this.toolTip1.SetToolTip(this.checkBoxLockTrees, "Kun valinta on päällä, puiden sisältö ei päivity mallista valittaessa");
            this.checkBoxLockTrees.UseVisualStyleBackColor = true;
            // 
            // toolStrip1
            // 
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripButtonCure});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(537, 25);
            this.toolStrip1.TabIndex = 3;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // toolStripButtonCure
            // 
            this.toolStripButtonCure.BackColor = System.Drawing.Color.Red;
            this.toolStripButtonCure.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripButtonCure.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButtonCure.Name = "toolStripButtonCure";
            this.toolStripButtonCure.Size = new System.Drawing.Size(123, 24);
            this.toolStripButtonCure.Text = "Korjaa kytkennät";
            this.toolStripButtonCure.ToolTipText = "Liitä ei-betoniosat alikokoonpanoina";
            this.toolStripButtonCure.Visible = false;
            this.toolStripButtonCure.Click += new System.EventHandler(this.toolStripButtonCure_Click);
            // 
            // AssemblyTree
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(537, 689);
            this.Controls.Add(this.toolStrip1);
            this.Controls.Add(this.checkBoxLockTrees);
            this.Controls.Add(this.treeView2);
            this.Controls.Add(this.treeView1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "AssemblyTree";
            this.Text = "Kokoonpanopuu";
            this.TopMost = true;
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form1_FormClosed);
            this.contextMenuStrip1.ResumeLayout(false);
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TreeView treeView1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem copyToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem pasteToolStripMenuItem;
        private System.Windows.Forms.TreeView treeView2;
        private System.Windows.Forms.CheckBox checkBoxLockTrees;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton toolStripButtonCure;
        private System.Windows.Forms.ToolStripMenuItem cutToolStripMenuItem;
    }

    public class TreeClipBoard
    {
        public TreeClipBoard()
        {
        }

        public TreeClipBoard(TreeOperationEnum operationEnum)
        {
            _operationEnum = operationEnum;
        }

        private Model _tsModel = new Model();
        private ModelObject _sourceModelObject = null;
        private Assembly _sourceRootAssembly = null;
        private TSG.CoordinateSystem _sourceCoordsys = null;
        private TSG.CoordinateSystem _targetCoordsys = null;
        private TreeOperationEnum _operationEnum = TreeOperationEnum.COPY;
        private bool _childrenOnly = false;

        public TreeOperationEnum CurrentOperation
        {
            get
            {
                return _operationEnum;
            }
        }

        public void AddObjectToClipBoard(Identifier selectedID)
        {
            _sourceModelObject = _tsModel.SelectModelObject(selectedID);
            Assembly fatherAssembly = null;

            if (_sourceModelObject is IAssemblable)
            {
                fatherAssembly = (_sourceModelObject as IAssemblable).GetAssembly();
                _sourceCoordsys = fatherAssembly.GetMainPart().GetCoordinateSystem();
            }
            else if (_sourceModelObject is Assembly)
            {
                fatherAssembly = _sourceModelObject as Assembly;
                _sourceCoordsys = fatherAssembly.GetAssembly().GetMainPart().GetCoordinateSystem();
            }
            else if (_sourceModelObject is Reinforcement)
            {
                fatherAssembly = ((_sourceModelObject as Reinforcement).Father as Part).GetAssembly();
                _sourceCoordsys = fatherAssembly.GetMainPart().GetCoordinateSystem();
            }

            _sourceRootAssembly = GetRootAssembly(fatherAssembly);
        }

        public void AddChildrenToClipBoard(Identifier selectedID)
        {
            _sourceModelObject = _tsModel.SelectModelObject(selectedID);
            _sourceCoordsys = _sourceModelObject.GetCoordinateSystem();
            _childrenOnly = true;
        }

        public void DumpClipBoardHere(Identifier selectedID)
        {
            ModelObject targetFather = _tsModel.SelectModelObject(selectedID);
            Assembly targetAssembly = null;

            if (targetFather is Assembly)
                targetAssembly = targetFather as Assembly;
            else if (targetFather is Part)
                targetAssembly = (targetFather as Part).GetAssembly();

            if (_sourceRootAssembly.Identifier.ID == GetRootAssembly(targetAssembly).Identifier.ID &&
                _operationEnum == TreeOperationEnum.CUT)
            {
                ChangeRelationToTarget(targetAssembly);
            }
            else
            {
                PasteToTarget(targetAssembly);
            }

            _tsModel.CommitChanges();
        }

        private void ChangeRelationToTarget(Assembly targetAssembly)
        {
            if (_sourceModelObject is Assembly)
            {
                Assembly sourceAssembly = (_sourceModelObject as Assembly).GetAssembly();

                sourceAssembly.Remove(_sourceModelObject);
                sourceAssembly.Modify();

                targetAssembly.Add(_sourceModelObject as Assembly);
                targetAssembly.Modify();
            }
            else if (_sourceModelObject is Part)
            {
                AssemblyTree.RemoveWeldsAndBolts(_sourceModelObject as Part);

                Assembly sourceAssembly = (_sourceModelObject as Part).GetAssembly();
                sourceAssembly.Remove(_sourceModelObject as Part);
                sourceAssembly.Modify();

                targetAssembly.Add(_sourceModelObject as Part);
                targetAssembly.Modify();
            }
        }

        private void PasteToTarget(Assembly targetAssembly)
        {
            _targetCoordsys = targetAssembly.GetMainPart().GetCoordinateSystem();

            if (_sourceModelObject is Assembly)
            {
                if (!_childrenOnly)
                {
                    CopyAssembly(_sourceModelObject as Assembly, targetAssembly);
                }
                else
                {
                    ArrayList subAssemblies = (_sourceModelObject as Assembly).GetSubAssemblies();

                    foreach (Assembly subAssembly in subAssemblies)
                        CopyAssembly(subAssembly, targetAssembly);

                    CopyChildren(_sourceModelObject as Assembly, targetAssembly);
                }
            }
            else
            {
                ModelObject resultObject = CopyObject(_sourceModelObject);

                if (resultObject is IAssemblable)
                {
                    targetAssembly.Add(resultObject as IAssemblable);
                    targetAssembly.Modify();
                }
            }
        }

        private Assembly GetRootAssembly(Assembly assembly)
        {
            Assembly result = assembly;
            Assembly fatherAssembly = assembly.GetAssembly();

            while (fatherAssembly != null)
            {
                result = fatherAssembly;
                fatherAssembly = fatherAssembly.GetAssembly();
            }

            return result;
        }

        private void CopyChildren(Assembly sourceAssembly, Assembly targetAssembly)
        {
            ArrayList secondaries = sourceAssembly.GetSecondaries();

            foreach(ModelObject sec in secondaries)
            {
                ModelObject resultObject = CopyObject(sec);

                if (resultObject is IAssemblable)
                    targetAssembly.Add(resultObject as IAssemblable);
            }

            targetAssembly.Modify();
        }

        private Assembly CopyAssembly(Assembly sourceAssembly, Assembly targetAssembly)
        {
            Assembly result = null;

            if (targetAssembly != null)
            {
                ModelObject resultMainPart = CopyObject(sourceAssembly.GetMainPart());
                result = (resultMainPart as Part).GetAssembly();

                ArrayList subAssemblies = sourceAssembly.GetSubAssemblies();

                foreach (Assembly subAssembly in subAssemblies)
                    CopyAssembly(subAssembly, result);

                CopyChildren(sourceAssembly, result);

                targetAssembly.Add(result);
                targetAssembly.Modify();

                if (_operationEnum == TreeOperationEnum.CUT)
                    sourceAssembly.Delete();
            }

            return result;
        }

        private ModelObject CopyObject(ModelObject sourceObject)
        {
            ModelObject result = Operation.CopyObject(sourceObject, _sourceCoordsys, _targetCoordsys);

            if (_operationEnum == TreeOperationEnum.CUT)
                sourceObject.Delete();

            return result;
        }
    }

    public enum TreeOperationEnum
    {
        CUT = 0,
        COPY = 1
    }

    public class Script
    {
        /// <summary>
        /// The method implementing A-Kit macro call which starts our dialog application.
        /// When this application is run from TS as a macro A-Kit commands will also
        /// be available.
        /// </summary>
        /// <param name="akit">the A-Kit object that Tekla Structures will pass on to the macro if run as a .NET macro</param>
        public static void Run(Tekla.Technology.Akit.IScript akit)
        {
            Application.EnableVisualStyles();
            Application.Run(new AssemblyTree());
        }
    }
}
