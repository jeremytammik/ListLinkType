#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Windows.Forms;
#endregion

namespace ListLinkType
{
  [Transaction( TransactionMode.ReadOnly )]
  public class Command : IExternalCommand
  {
    /// <summary>
    /// Indent a string by adding padding left.
    /// </summary>
    static string Indent( int count )
    {
      return "".PadLeft( count );
    }

    /// <summary>
    /// Return a label to display for a link type.
    /// Nested links are always attachments.
    /// </summary>
    static string LinkLabel( 
      RevitLinkType type, 
      int level )
    {
      AttachmentType a = type.AttachmentType;
      AttachmentType b = 0 < level
        ? AttachmentType.Attachment
        : a;

      Debug.Print(
        "AttachmentType {0} at level {1} -> {2}", 
        a, level, b );

      return type.Name + "->" + b.ToString();
    }

    /// <summary>
    /// Recursively retrieve entire linked document 
    /// hierarchy and return the resulting TreeNode
    /// structure.
    /// </summary>
    void GetChildren( 
      Document mainDoc, 
      ICollection<ElementId> ids, 
      TreeNode parentNode )
    {
      int level = parentNode.Level;

      foreach( ElementId id in ids )
      {
        // Get the child information.

        RevitLinkType type = mainDoc.GetElement( id ) 
          as RevitLinkType;

        string label = LinkLabel( type, level );

        TreeNode subNode = new TreeNode( label );

        Debug.Print( "{0}{1}", Indent( 2 * level ), 
          label );

        parentNode.Nodes.Add( subNode );

        // Go to the next level.

        GetChildren( mainDoc, type.GetChildIds(), 
          subNode );
      }
    }

    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements )
    {
      UIApplication uiapp = commandData.Application;
      UIDocument uidoc = uiapp.ActiveUIDocument;

      // Get the active document.

      Document mainDoc = uidoc.Document;

      // Prepare to show the resulting linked 
      // document tree hierarchy.

      TreeNode mainNode = new TreeNode();
      mainNode.Text = mainDoc.PathName;

      // Start at the root links (no parent node).

      FilteredElementCollector coll 
        = new FilteredElementCollector( mainDoc )
          .OfClass( typeof( RevitLinkInstance ) );

      foreach( RevitLinkInstance inst in coll )
      {
        RevitLinkType type = mainDoc.GetElement( 
          inst.GetTypeId() ) as RevitLinkType;

        if( type.GetParentId() 
          == ElementId.InvalidElementId )
        {
          string label = LinkLabel( type, 0 );

          TreeNode parentNode = new TreeNode( label );

          Debug.Print( label );

          mainNode.Nodes.Add( parentNode );

          GetChildren( mainDoc, type.GetChildIds(), 
            parentNode );
        }
      }

      // Show the results in a form.

      System.Windows.Forms.Form resultForm 
        = new System.Windows.Forms.Form();

      TreeView treeView = new TreeView();
      treeView.Size = resultForm.Size;
      treeView.Anchor |= AnchorStyles.Bottom 
        | AnchorStyles.Top;

      treeView.Nodes.Add( mainNode );
      resultForm.Controls.Add( treeView );
      resultForm.ShowDialog();

      return Result.Succeeded;
    }
  }
}
