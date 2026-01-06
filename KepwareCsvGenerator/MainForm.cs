using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Windows.Forms;

namespace KepwareCsvGenerator;

public sealed class MainForm : Form
{
    private readonly TreeView _treeView;
    private readonly Button _addFolderButton;
    private readonly Button _addTagButton;
    private readonly Button _duplicateButton;
    private readonly Button _exportButton;
    private readonly Button _saveButton;
    private readonly Button _loadButton;
    private readonly Button _removeButton;
    private readonly ContextMenuStrip _contextMenu = new();
    private readonly ToolStripMenuItem _duplicateMenuItem = new("Duplicate Folder");
    private TreeNode? _dragNode;

    public MainForm()
    {
        Text = "Kepware CSV Generator";
        Width = 900;
        Height = 600;
        StartPosition = FormStartPosition.CenterScreen;
        AutoScaleMode = AutoScaleMode.Dpi;
        AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);

        var headerLabel = new Label
        {
            Text = "Arrange folders/tags, then export Kepware CSV",
            Dock = DockStyle.Top,
            Height = 30,
            TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        };

        _treeView = new TreeView
        {
            Dock = DockStyle.Fill,
            HideSelection = false,
            AllowDrop = true
        };

        _treeView.ItemDrag += TreeViewOnItemDrag;
        _treeView.DragEnter += TreeViewOnDragEnter;
        _treeView.DragDrop += TreeViewOnDragDrop;
        _treeView.NodeMouseClick += TreeViewOnNodeMouseClick;
        _treeView.NodeMouseDoubleClick += TreeViewOnNodeMouseDoubleClick;
        ConfigureContextMenu();

        _addFolderButton = new Button
        {
            Text = "Add Folder",
            Width = 120
        };
        StyleActionButton(_addFolderButton);
        _addFolderButton.Click += (_, _) => AddFolder();

        _addTagButton = new Button
        {
            Text = "Add Tag",
            Width = 120
        };
        StyleActionButton(_addTagButton);
        _addTagButton.Click += (_, _) => AddTag();

        _duplicateButton = new Button
        {
            Text = "Duplicate Folder",
            Width = 120
        };
        StyleActionButton(_duplicateButton);
        _duplicateButton.Click += (_, _) => DuplicateSelectedFolder();

        _removeButton = new Button
        {
            Text = "Remove",
            Width = 120
        };
        StyleActionButton(_removeButton);
        _removeButton.Click += (_, _) => RemoveSelected();

        _exportButton = new Button
        {
            Text = "Export CSV",
            Width = 120
        };
        StyleActionButton(_exportButton);
        _exportButton.Click += (_, _) => ExportCsv();

        _saveButton = new Button
        {
            Text = "Save JSON",
            Width = 120
        };
        StyleActionButton(_saveButton);
        _saveButton.Click += (_, _) => SaveHierarchy();

        _loadButton = new Button
        {
            Text = "Load JSON",
            Width = 120
        };
        StyleActionButton(_loadButton);
        _loadButton.Click += (_, _) => LoadHierarchy();

        var buttonPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            ColumnCount = 1,
            RowCount = 7,
            Padding = new Padding(10)
        };
        buttonPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        buttonPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        buttonPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        buttonPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        buttonPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        buttonPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        buttonPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        buttonPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        buttonPanel.Controls.Add(_addFolderButton, 0, 0);
        buttonPanel.Controls.Add(_addTagButton, 0, 1);
        buttonPanel.Controls.Add(_duplicateButton, 0, 2);
        buttonPanel.Controls.Add(_removeButton, 0, 3);
        buttonPanel.Controls.Add(_saveButton, 0, 4);
        buttonPanel.Controls.Add(_loadButton, 0, 5);
        buttonPanel.Controls.Add(_exportButton, 0, 6);

        var rightPanel = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(0, 10, 10, 10)
        };
        rightPanel.Controls.Add(buttonPanel);

        var mainLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 1
        };
        mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70));
        mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
        mainLayout.Controls.Add(_treeView, 0, 0);
        mainLayout.Controls.Add(rightPanel, 1, 0);

        Controls.Add(mainLayout);
        Controls.Add(headerLabel);

        InitializeRoot();
    }

    private void InitializeRoot()
    {
        var root = new TreeNode("Root")
        {
            Tag = NodeMetadata.Folder("Root")
        };
        _treeView.Nodes.Add(root);
        _treeView.SelectedNode = root;
        root.Expand();
    }

    private void ConfigureContextMenu()
    {
        _duplicateMenuItem.Click += (_, _) => DuplicateSelectedFolder();
        _contextMenu.Items.Add(_duplicateMenuItem);
        _contextMenu.Opening += (_, e) =>
        {
            var selected = _treeView.SelectedNode;
            var enabled = selected?.Tag is NodeMetadata { IsFolder: true } && selected.Parent != null;
            _duplicateMenuItem.Enabled = enabled;
            e.Cancel = !enabled;
        };

        _treeView.ContextMenuStrip = _contextMenu;
    }

    private void AddFolder()
    {
        var parent = GetSelectedContainerNode();
        if (parent == null)
        {
            MessageBox.Show("Select a folder to add a new folder.", "Invalid Selection", MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        using var dialog = new NameDialog("New Folder", "Folder name:");
        if (dialog.ShowDialog(this) != DialogResult.OK || string.IsNullOrWhiteSpace(dialog.EntityName))
        {
            return;
        }

        var node = new TreeNode(dialog.EntityName)
        {
            Tag = NodeMetadata.Folder(dialog.EntityName)
        };
        parent.Nodes.Add(node);
        parent.Expand();
        _treeView.SelectedNode = node;
    }

    private void AddTag()
    {
        var parent = GetSelectedContainerNode();
        if (parent == null)
        {
            MessageBox.Show("Select a folder to add a tag.", "Invalid Selection", MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        using var dialog = new TagDialog();
        if (dialog.ShowDialog(this) != DialogResult.OK || string.IsNullOrWhiteSpace(dialog.TagName))
        {
            return;
        }

        var metadata = NodeMetadata.Tag(dialog.TagName, dialog.SelectedDataType);
        var node = new TreeNode(dialog.TagName)
        {
            Tag = metadata
        };

        parent.Nodes.Add(node);
        parent.Expand();
        _treeView.SelectedNode = node;
    }

    private void DuplicateSelectedFolder()
    {
        var selected = _treeView.SelectedNode;
        if (selected?.Tag is not NodeMetadata { IsFolder: true } || selected.Parent == null)
        {
            MessageBox.Show("Select a folder (not the root) to duplicate.", "Invalid Selection", MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        var newName = GenerateUniqueFolderName(selected.Text, selected.Parent.Nodes);
        var clone = CloneSubtree(selected, newName);
        selected.Parent.Nodes.Add(clone);
        selected.Parent.Expand();
        _treeView.SelectedNode = clone;
    }

    private void RemoveSelected()
    {
        var node = _treeView.SelectedNode;
        if (node == null || node.Parent == null)
        {
            return;
        }

        node.Remove();
    }

    private void ExportCsv()
    {
        var tags = CollectTags();
        if (tags.Count == 0)
        {
            MessageBox.Show("No tags found. Add tags before exporting.", "No Tags", MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            return;
        }

        using var saveDialog = new SaveFileDialog
        {
            Filter = "CSV Files (*.csv)|*.csv",
            FileName = "kepware_tags.csv"
        };

        if (saveDialog.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        var csv = CsvBuilder.Build(tags);
        File.WriteAllText(saveDialog.FileName, csv);
        MessageBox.Show("CSV exported successfully.", "Export Complete", MessageBoxButtons.OK,
            MessageBoxIcon.Information);
    }

    private void SaveHierarchy()
    {
        var root = _treeView.Nodes[0];
        var model = HierarchyNode.FromTreeNode(root);
        using var saveDialog = new SaveFileDialog
        {
            Filter = "JSON Files (*.json)|*.json",
            FileName = "kepware_hierarchy.json"
        };

        if (saveDialog.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        var json = JsonSerializer.Serialize(model, JsonSerializerOptionsFactory.Create());
        File.WriteAllText(saveDialog.FileName, json);
        MessageBox.Show("Hierarchy saved.", "Save Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void LoadHierarchy()
    {
        using var openDialog = new OpenFileDialog
        {
            Filter = "JSON Files (*.json)|*.json"
        };

        if (openDialog.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        var json = File.ReadAllText(openDialog.FileName);
        var model = JsonSerializer.Deserialize<HierarchyNode>(json, JsonSerializerOptionsFactory.Create());
        if (model == null)
        {
            MessageBox.Show("Invalid JSON file.", "Load Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        _treeView.Nodes.Clear();
        var rootNode = model.ToTreeNode();
        _treeView.Nodes.Add(rootNode);
        _treeView.SelectedNode = rootNode;
        rootNode.Expand();
    }

    private TreeNode? GetSelectedContainerNode()
    {
        var selected = _treeView.SelectedNode ?? _treeView.Nodes[0];
        var metadata = selected.Tag as NodeMetadata;
        if (metadata is { IsFolder: true })
        {
            return selected;
        }

        return selected.Parent;
    }

    private List<TagRecord> CollectTags()
    {
        var tags = new List<TagRecord>();
        var root = _treeView.Nodes[0];
        foreach (TreeNode node in root.Nodes)
        {
            TraverseNode(node, new List<string>(), tags);
        }

        return tags;
    }

    private static void TraverseNode(TreeNode node, List<string> pathSegments, List<TagRecord> output)
    {
        if (node.Tag is not NodeMetadata metadata)
        {
            return;
        }

        if (metadata.IsFolder)
        {
            pathSegments.Add(metadata.Name);
            foreach (TreeNode child in node.Nodes)
            {
                TraverseNode(child, pathSegments, output);
            }
            pathSegments.RemoveAt(pathSegments.Count - 1);
            return;
        }

        var fullName = string.Join('.', pathSegments.Append(metadata.Name));
        output.Add(new TagRecord(fullName, metadata.DataType));
    }

    private static string GenerateUniqueFolderName(string baseName, TreeNodeCollection siblings)
    {
        var index = 1;
        string candidate;
        do
        {
            candidate = $"{baseName} ({index})";
            index++;
        } while (siblings.Cast<TreeNode>().Any(node =>
                     string.Equals(node.Text, candidate, StringComparison.OrdinalIgnoreCase)));

        return candidate;
    }

    private static TreeNode CloneSubtree(TreeNode source, string? rootNameOverride = null)
    {
        if (source.Tag is not NodeMetadata metadata)
        {
            return (TreeNode)source.Clone();
        }

        var name = rootNameOverride ?? metadata.Name;
        var cloneMetadata = metadata.IsFolder
            ? NodeMetadata.Folder(name)
            : NodeMetadata.Tag(name, metadata.DataType);

        var clone = new TreeNode(name)
        {
            Tag = cloneMetadata
        };

        foreach (TreeNode child in source.Nodes)
        {
            clone.Nodes.Add(CloneSubtree(child));
        }

        return clone;
    }

    private static void StyleActionButton(Button button)
    {
        button.AutoSize = true;
        button.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        button.MinimumSize = new Size(120, 36);
        button.Padding = new Padding(6, 8, 6, 8);
        button.Anchor = AnchorStyles.Left | AnchorStyles.Right;
    }

    private void TreeViewOnItemDrag(object? sender, ItemDragEventArgs e)
    {
        if (e.Item is TreeNode node && node.Parent != null)
        {
            _dragNode = node;
            DoDragDrop(e.Item, DragDropEffects.Move);
        }
    }

    private void TreeViewOnDragEnter(object? sender, DragEventArgs e)
    {
        e.Effect = DragDropEffects.Move;
    }

    private void TreeViewOnDragDrop(object? sender, DragEventArgs e)
    {
        if (_dragNode == null)
        {
            return;
        }

        var targetPoint = _treeView.PointToClient(new System.Drawing.Point(e.X, e.Y));
        var targetNode = _treeView.GetNodeAt(targetPoint);
        if (targetNode == null)
        {
            return;
        }

        var targetMetadata = targetNode.Tag as NodeMetadata;
        if (targetMetadata is not { IsFolder: true })
        {
            targetNode = targetNode.Parent;
        }

        if (targetNode == null || _dragNode == targetNode || IsDescendantOf(targetNode, _dragNode))
        {
            return;
        }

        _dragNode.Remove();
        targetNode.Nodes.Add(_dragNode);
        targetNode.Expand();
        _treeView.SelectedNode = _dragNode;
    }

    private static bool IsDescendantOf(TreeNode parent, TreeNode candidate)
    {
        var current = parent;
        while (current.Parent != null)
        {
            if (current.Parent == candidate)
            {
                return true;
            }
            current = current.Parent;
        }

        return false;
    }

    private void TreeViewOnNodeMouseClick(object? sender, TreeNodeMouseClickEventArgs e)
    {
        _treeView.SelectedNode = e.Node;
    }

    private void TreeViewOnNodeMouseDoubleClick(object? sender, TreeNodeMouseClickEventArgs e)
    {
        _treeView.SelectedNode = e.Node;
        if (e.Node?.Tag is not NodeMetadata metadata)
        {
            return;
        }

        var isFolder = metadata.IsFolder;
        var dialogTitle = isFolder ? "Rename Folder" : "Rename Tag";
        var dialogLabel = isFolder ? "Folder name:" : "Tag name:";
        using var dialog = new NameDialog(dialogTitle, dialogLabel, metadata.Name);
        if (dialog.ShowDialog(this) != DialogResult.OK || string.IsNullOrWhiteSpace(dialog.EntityName))
        {
            return;
        }

        var newName = dialog.EntityName;
        e.Node.Text = newName;
        e.Node.Tag = isFolder
            ? NodeMetadata.Folder(newName)
            : NodeMetadata.Tag(newName, metadata.DataType);
    }
}

public sealed class NodeMetadata
{
    public string Name { get; }
    public bool IsFolder { get; }
    public TagDataType DataType { get; }

    private NodeMetadata(string name, bool isFolder, TagDataType dataType)
    {
        Name = name;
        IsFolder = isFolder;
        DataType = dataType;
    }

    public static NodeMetadata Folder(string name) => new(name, true, TagDataType.String);

    public static NodeMetadata Tag(string name, TagDataType dataType) => new(name, false, dataType);
}

public enum TagDataType
{
    String,
    Integer,
    Boolean
}

public sealed record TagRecord(string Name, TagDataType DataType);

public static class CsvBuilder
{
    private static readonly string Header =
        "Tag Name,Address,Data Type,Respect Data Type,Client Access,Scan Rate,Scaling,Raw Low,Raw High,Scaled Low,Scaled High,Scaled Data Type,Clamp Low,Clamp High,Eng Units,Description,Negate Value";

    public static string Build(IReadOnlyList<TagRecord> tags)
    {
        var addressGenerator = new AddressGenerator();
        var builder = new StringBuilder();
        builder.AppendLine(Header);

        foreach (var tag in tags)
        {
            var address = addressGenerator.Next(tag.DataType);
            var dataTypeString = tag.DataType switch
            {
                TagDataType.String => "string",
                TagDataType.Integer => "integer",
                TagDataType.Boolean => "boolean",
                _ => "string"
            };

            builder.Append(tag.Name);
            builder.Append(',');
            builder.Append(address);
            builder.Append(',');
            builder.Append(dataTypeString);
            builder.Append(",1,R/W,100");
            builder.AppendLine(",,,,,,,,,,,");
        }

        return builder.ToString();
    }
}

public sealed class AddressGenerator
{
    private int _stringIndex = 1;
    private int _integerIndex = 0;
    private int _booleanWordIndex = 0;
    private int _booleanBitIndex = 0;

    public string Next(TagDataType dataType)
    {
        return dataType switch
        {
            TagDataType.String => $"S{_stringIndex++:000}",
            TagDataType.Integer => $"D{_integerIndex++:0000}",
            TagDataType.Boolean => NextBoolean(),
            _ => $"S{_stringIndex++:000}"
        };
    }

    private string NextBoolean()
    {
        var address = $"D{_booleanWordIndex:0000}.{_booleanBitIndex}";
        _booleanBitIndex++;
        if (_booleanBitIndex >= 16)
        {
            _booleanBitIndex = 0;
            _booleanWordIndex++;
        }

        return address;
    }
}

public sealed class NameDialog : Form
{
    private readonly TextBox _nameBox;

    public NameDialog(string title, string label, string initialText = "")
    {
        Text = title;
        Width = 360;
        Height = 150;
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;

        var labelControl = new Label
        {
            Text = label,
            Dock = DockStyle.Top,
            Height = 24
        };

        _nameBox = new TextBox
        {
            Dock = DockStyle.Top,
            Text = initialText
        };

        var buttonPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            FlowDirection = FlowDirection.RightToLeft,
            Height = 40
        };

        var okButton = new Button { Text = "OK", DialogResult = DialogResult.OK };
        var cancelButton = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel };
        buttonPanel.Controls.Add(okButton);
        buttonPanel.Controls.Add(cancelButton);

        Controls.Add(buttonPanel);
        Controls.Add(_nameBox);
        Controls.Add(labelControl);

        AcceptButton = okButton;
        CancelButton = cancelButton;
    }

    public string EntityName => _nameBox.Text.Trim();
}

public sealed class TagDialog : Form
{
    private readonly TextBox _nameBox;
    private readonly ComboBox _dataTypeBox;

    public TagDialog()
    {
        Text = "New Tag";
        Width = 360;
        Height = 200;
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;

        var nameLabel = new Label
        {
            Text = "Tag name:",
            Dock = DockStyle.Top,
            Height = 24
        };

        _nameBox = new TextBox
        {
            Dock = DockStyle.Top
        };

        var dataTypeLabel = new Label
        {
            Text = "Data type:",
            Dock = DockStyle.Top,
            Height = 24
        };

        _dataTypeBox = new ComboBox
        {
            Dock = DockStyle.Top,
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        _dataTypeBox.Items.AddRange(Enum.GetNames(typeof(TagDataType)));
        _dataTypeBox.SelectedIndex = 0;

        var buttonPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            FlowDirection = FlowDirection.RightToLeft,
            Height = 40
        };

        var okButton = new Button { Text = "OK", DialogResult = DialogResult.OK };
        var cancelButton = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel };
        buttonPanel.Controls.Add(okButton);
        buttonPanel.Controls.Add(cancelButton);

        Controls.Add(buttonPanel);
        Controls.Add(_dataTypeBox);
        Controls.Add(dataTypeLabel);
        Controls.Add(_nameBox);
        Controls.Add(nameLabel);

        AcceptButton = okButton;
        CancelButton = cancelButton;
    }

    public string TagName => _nameBox.Text.Trim();

    public TagDataType SelectedDataType => Enum.TryParse<TagDataType>(_dataTypeBox.SelectedItem?.ToString(), out var value)
        ? value
        : TagDataType.String;
}

public sealed class HierarchyNode
{
    public string Name { get; init; } = string.Empty;
    public bool IsFolder { get; init; }
    public TagDataType DataType { get; init; }
    public List<HierarchyNode> Children { get; init; } = new();

    public static HierarchyNode FromTreeNode(TreeNode node)
    {
        var metadata = node.Tag as NodeMetadata;
        var model = new HierarchyNode
        {
            Name = metadata?.Name ?? node.Text,
            IsFolder = metadata?.IsFolder ?? true,
            DataType = metadata?.DataType ?? TagDataType.String
        };

        foreach (TreeNode child in node.Nodes)
        {
            model.Children.Add(FromTreeNode(child));
        }

        return model;
    }

    public TreeNode ToTreeNode()
    {
        var metadata = IsFolder ? NodeMetadata.Folder(Name) : NodeMetadata.Tag(Name, DataType);
        var node = new TreeNode(Name)
        {
            Tag = metadata
        };

        foreach (var child in Children)
        {
            node.Nodes.Add(child.ToTreeNode());
        }

        return node;
    }
}

public static class JsonSerializerOptionsFactory
{
    public static JsonSerializerOptions Create()
    {
        return new JsonSerializerOptions
        {
            WriteIndented = true
        };
    }
}
