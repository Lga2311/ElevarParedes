Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry

Public Class Class1

    <CommandMethod("DrawLine")>
    Public Sub DrawLine()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim edt As Editor = doc.Editor

        Using trans As Transaction = db.TransactionManager.StartTransaction()

            Try
                Dim bt As BlockTable
                bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead)

                Dim btr As BlockTableRecord
                btr = trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                edt.WriteMessage("Drawing a Line object" + vbCrLf)
                Dim pt1 As Point3d = New Point3d(0, 0, 0)
                Dim pt2 As Point3d = New Point3d(25, 150, 0)
                Dim ln As Line = New Line(pt1, pt2)
                ln.ColorIndex = 1
                btr.AppendEntity(ln)
                trans.AddNewlyCreatedDBObject(ln, True)
                trans.Commit()

            Catch ex As Exception
                edt.WriteMessage("Error encountered: " + ex.Message)
                trans.Abort()

            End Try

        End Using

    End Sub

End Class
