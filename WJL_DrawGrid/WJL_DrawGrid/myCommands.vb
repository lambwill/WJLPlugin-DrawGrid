' (C) Copyright 2011 by  
'
Imports System
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

' This line is not mandatory, but improves loading performances
<Assembly: CommandClass(GetType(WJL_DrawGrid.DrawGrid))> 
Namespace WJL_DrawGrid
    Public Class DrawGrid
        Dim PersistantSettings As New GridSettings   'Instance of "GridSettings" to store persistant settings

        Enum LabelAxisEnum As Integer
            glNone = 0
            glTop = 1
            glRight = 2
            glBottom = 4
            glLeft = 8
        End Enum

        'Class to hold settings....
        Public Class GridSettings

            'Set defaults 
            Private mySetDefaults As Boolean = True
            Property SetDefaults() As Boolean
                Get
                    SetDefaults = mySetDefaults
                End Get
                Set(ByVal value As Boolean)
                    mySetDefaults = value
                End Set
            End Property

            'BlockID for crosshair block
            Private myBlockID As ObjectId
            Property BlockID() As ObjectId
                Get
                    BlockID = myBlockID
                End Get
                Set(ByVal value As ObjectId)
                    myBlockID = value
                End Set
            End Property

            'Number of points in X direction
            Private myNumberOfX As Integer
            Property NumberOfX() As Integer
                Get
                    NumberOfX = myNumberOfX
                End Get
                Set(ByVal value As Integer)
                    myNumberOfX = value
                End Set
            End Property

            'Number of points in Y direction
            Private myNumberOfY As Integer
            Property NumberOfY() As Integer
                Get
                    NumberOfY = myNumberOfY
                End Get
                Set(ByVal value As Integer)
                    myNumberOfY = value
                End Set
            End Property

            'Base point for grid
            Private myBasePoint As Point3d
            Property BasePoint() As Point3d
                Get
                    BasePoint = myBasePoint
                End Get
                Set(ByVal value As Point3d)
                    myBasePoint = value
                End Set
            End Property

            'Local / global grid
            Private myLocalGrid As Boolean
            Property LocalGrid() As Boolean
                Get
                    LocalGrid = myLocalGrid
                End Get
                Set(ByVal value As Boolean)
                    myLocalGrid = value
                End Set
            End Property

            'Extents of grid
            Private myGridExtents As Point3d
            Property GridExtents() As Point3d
                Get
                    GridExtents = myGridExtents
                End Get
                Set(ByVal value As Point3d)
                    myGridExtents = value
                End Set
            End Property

            'Size of grid crosshairs
            Private myCrosshairSize As Double
            Property CrosshairSize() As Double
                Get
                    CrosshairSize = myCrosshairSize
                End Get
                Set(ByVal value As Double)
                    myCrosshairSize = value
                End Set
            End Property

            'Interval between grid lines
            Private myGridInterval As Double
            Property GridInterval() As Double
                Get
                    GridInterval = myGridInterval
                End Get
                Set(ByVal value As Double)
                    myGridInterval = value
                End Set
            End Property


        End Class

        'DrawJig class to handle input jig
        Public Class InputJig
            Inherits Autodesk.AutoCAD.EditorInput.DrawJig
            Private myJigSettings As GridSettings
            Dim myPR As PromptPointResult

            Private myKeyword As String
            Property Keyword() As String
                Get
                    Return myKeyword
                End Get
                Set(ByVal value As String)
                    myKeyword = value
                End Set
            End Property

            Property JigSettings() As GridSettings
                Get
                    Return myJigSettings
                End Get
                Set(ByVal value As GridSettings)
                    myJigSettings = value
                End Set
            End Property

            Function StartJig(ByVal BasePt As Point3d, ByRef Settings As GridSettings) As PromptResult 'PromptPointResult
                myJigSettings = Settings

                Dim myED As Editor = Application.DocumentManager.MdiActiveDocument.Editor
                myPR = myED.Drag(Me)

                Do
                    Select Case myPR.Status
                        Case PromptStatus.None
                            Return myPR
                            Exit Do
                        Case PromptStatus.OK
                            Return myPR
                            Exit Do
                        Case PromptStatus.Keyword
                            Return myPR
                            Exit Do
                    End Select
                Loop While myPR.Status <> PromptStatus.Cancel

                Return myPR

            End Function

            Protected Overrides Function Sampler(ByVal prompts As Autodesk.AutoCAD.EditorInput.JigPrompts) As Autodesk.AutoCAD.EditorInput.SamplerStatus
                Dim GetStringOptions As New JigPromptPointOptions

                Dim PromptString As String = vbLf & "Select grid extents or "
                PromptString = PromptString & "[grid Interval/crosshair Size/Coordinate system] : "
                Dim Keywords As String = "Interval Size Coordinates"

                Dim UserPoint As Point3d
                Dim BasePoint As Point3d = JigSettings.BasePoint

                GetStringOptions.SetMessageAndKeywords(PromptString, Keywords)
                GetStringOptions.AppendKeywordsToMessage = True
                GetStringOptions.Keywords.GetDisplayString(True)
                GetStringOptions.UserInputControls = UserInputControls.NullResponseAccepted

                Dim JigPropmtStringResult As PromptPointResult
                JigPropmtStringResult = prompts.AcquirePoint(GetStringOptions)

                myKeyword = JigPropmtStringResult.StringResult

                Select Case JigPropmtStringResult.Status
                    Case PromptStatus.None
                        Return SamplerStatus.OK

                    Case PromptStatus.Keyword
                        Return SamplerStatus.OK

                    Case PromptStatus.OK
                        UserPoint = JigPropmtStringResult.Value
                        If UserPoint.X > BasePoint.X And UserPoint.Y > BasePoint.Y Then
                            JigSettings.GridExtents = UserPoint
                        Else
                            JigSettings.GridExtents = BasePoint
                        End If
                        Return SamplerStatus.OK

                    Case Else
                        Return SamplerStatus.Cancel

                End Select
            End Function

            Protected Overrides Function WorldDraw(ByVal draw As Autodesk.AutoCAD.GraphicsInterface.WorldDraw) As Boolean

                Dim GridWidth As Double = JigSettings.GridExtents.X - JigSettings.BasePoint.X
                Dim GridHeight As Double = JigSettings.GridExtents.Y - JigSettings.BasePoint.Y
                Dim GridInterval As Double = JigSettings.GridInterval

                Dim CrosshairCollection As DBObjectCollection = DrawJigCrossHairs(JigSettings, True)
                For i = 0 To CrosshairCollection.Count - 1
                    draw.Geometry.Draw(CrosshairCollection(i))
                Next

                For Each EntToDispose As Object In CrosshairCollection
                    EntToDispose.Dispose()
                Next
                CrosshairCollection.Dispose()

                Dim GridLabelCollection As DBObjectCollection = DrawGridLabels(JigSettings, True)
                For i = 0 To GridLabelCollection.Count - 1
                    draw.Geometry.Draw(GridLabelCollection(i))
                Next
                For Each EntToDispose As Object In GridLabelCollection
                    EntToDispose.Dispose()
                Next
                GridLabelCollection.Dispose()

            End Function

        End Class

        'Command DRAWGRID
        <CommandMethod("DRAWGRID")> _
        Public Sub DRAWGRID()

            'Restore GridSettings from last run (PersistantSettings)
            Dim myGridSettings As GridSettings = PersistantSettings

            If PersistantSettings.SetDefaults = True Then
                myGridSettings.GridInterval = 20
                myGridSettings.CrosshairSize = 5
                myGridSettings.LocalGrid = False

                PersistantSettings.SetDefaults = False
            End If

            Dim myDWG As Document = Application.DocumentManager.MdiActiveDocument
            Dim myDB As Database = myDWG.Database
            Dim myEd As Editor = myDWG.Editor

            'Get the grid basepoint
            Dim BasePointOptions As New PromptPointOptions("Select grid base point:")
            Dim BasePointResult As PromptPointResult = myEd.GetPoint(BasePointOptions)
            myGridSettings.BasePoint = BasePointResult.Value

            'set up the polylines for the croshair block 
            Dim CrosshairCollection As DBObjectCollection
            CrosshairCollection = CreateCrosshair(myGridSettings)

            'create (or aquire) the crosshair block
            Dim CrossHairBtrId As ObjectId
            CrossHairBtrId = CreateBlock(CrosshairCollection, "CrossHairBlock", myGridSettings.BasePoint)
            myGridSettings.BlockID = CrossHairBtrId
            'determine scale for crosshairs
            Dim CrossHairSize As Double = myGridSettings.CrosshairSize

RETRY:
            Dim CoorddinateSystem As String
            If myGridSettings.LocalGrid = True Then
                CoorddinateSystem = "local grid"
            Else
                CoorddinateSystem = "global grid"
            End If

            myEd.WriteMessage(vbLf & "Grid interval = " & myGridSettings.GridInterval & ",  Coordinate system = " & CoorddinateSystem & ", Crosshair size = " & myGridSettings.CrosshairSize)

            Dim myJig As InputJig = New InputJig()
            Dim GetPointResult As PromptResult
            GetPointResult = myJig.StartJig(myGridSettings.BasePoint, myGridSettings)
            Select Case GetPointResult.Status
                Case PromptStatus.OK
                    'Insert the crosshair blocks & axis lables
                    Using myTrans As Transaction = myDB.TransactionManager.StartTransaction()
                        ' Get the block table from the drawing
                        Dim bt As BlockTable = DirectCast(myTrans.GetObject(myDB.BlockTableId, OpenMode.ForRead), BlockTable)
                        ' get model space
                        Dim ms As BlockTableRecord = DirectCast(myTrans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)
                        'add crosshairs
                        Dim GridWidth As Double = myGridSettings.GridExtents.X - myGridSettings.BasePoint.X
                        Dim GridHeight As Double = myGridSettings.GridExtents.Y - myGridSettings.BasePoint.Y

                        Dim NumberOfX As Integer
                        NumberOfX = Math.Round(GridWidth / myGridSettings.GridInterval)

                        Dim NumberOfY As Integer
                        NumberOfY = Math.Round(GridHeight / myGridSettings.GridInterval)

                        Dim GridLabelCollection As DBObjectCollection = DrawGridLabels(myGridSettings)
                        For i = 0 To GridLabelCollection.Count - 1
                            ms.AppendEntity(GridLabelCollection(i))
                            myTrans.AddNewlyCreatedDBObject(GridLabelCollection(i), True)
                        Next

                        Dim GridCrosshairCollection As DBObjectCollection = DrawJigCrossHairs(myGridSettings)
                        For i = 0 To GridCrosshairCollection.Count - 1
                            ms.AppendEntity(GridCrosshairCollection(i))
                            myTrans.AddNewlyCreatedDBObject(GridCrosshairCollection(i), True)
                        Next
                        myTrans.Commit()
                        GridLabelCollection.Dispose()
                        GridCrosshairCollection.Dispose()
                    End Using

                Case PromptStatus.Keyword
                    Select Case myJig.Keyword 'GetPointResult.StringResult
                        Case "Interval"
                            Dim GridIntervalOptions As New PromptDoubleOptions("Specify grid interval: ")
                            GridIntervalOptions.AllowZero = False
                            GridIntervalOptions.AllowNegative = False
                            GridIntervalOptions.DefaultValue = myGridSettings.GridInterval
                            GridIntervalOptions.UseDefaultValue = True
                            Dim GridIntervalResult As PromptDoubleResult = myEd.GetDouble(GridIntervalOptions)
                            If GridIntervalResult.Status = PromptStatus.OK Then
                                myGridSettings.GridInterval = GridIntervalResult.Value
                            End If

                        Case "Size"
                            Dim GridCrosshairOptions As New PromptDoubleOptions("Specify grid crosshair size: ")
                            GridCrosshairOptions.AllowZero = True
                            GridCrosshairOptions.AllowNegative = False
                            GridCrosshairOptions.DefaultValue = myGridSettings.CrosshairSize
                            GridCrosshairOptions.UseDefaultValue = True
                            Dim GridCrosshairResult As PromptDoubleResult = myEd.GetDouble(GridCrosshairOptions)
                            If GridCrosshairResult.Status = PromptStatus.OK Then
                                myGridSettings.CrosshairSize = GridCrosshairResult.Value
                            End If

                        Case "Coordinates"
                            Dim PromptString As String = vbLf & "Specify coordinate system "
                            PromptString = PromptString & "[Local/Global] : "
                            Dim Keywords As String = "Local Global"

                            Dim CoordinatesOptions As New PromptKeywordOptions(PromptString, Keywords)
                            CoordinatesOptions.AppendKeywordsToMessage = True

                            Dim JigPropmtStringResult As PromptResult
                            JigPropmtStringResult = myEd.GetKeywords(CoordinatesOptions)

                            Select Case JigPropmtStringResult.StringResult
                                Case "Local"
                                    myGridSettings.LocalGrid = True
                                Case "Global"
                                    myGridSettings.LocalGrid = False
                            End Select

                    End Select
                    GoTo RETRY
                Case PromptStatus.Cancel
                    myEd.WriteMessage("*cancel*")
            End Select

SAVE_SETTINGS:
            PersistantSettings = myGridSettings
EXIT_WITHOUT_SAVE:
            'myEd.WriteMessage(vbLf & "DRAWGRID written by Will Lamb william.lamb@aecom.com")
        End Sub

        Function CreateCrosshair(ByVal myGridSettings As GridSettings) As DBObjectCollection
            Dim CrossLength As Double = 0.5

            Dim CrosshairCollection As New DBObjectCollection

            'define points for vertical line
            Dim VerticalLinePoints As New Point2dCollection
            VerticalLinePoints.Add(New Point2d(0, 0 - CrossLength))
            VerticalLinePoints.Add(New Point2d(0, 0 + CrossLength))

            'define vertical line
            Dim VerticalLine As New Polyline
            VerticalLine.AddVertexAt(0, VerticalLinePoints(0), 0, 0, 0)
            VerticalLine.AddVertexAt(1, VerticalLinePoints(1), 0, 0, 0)
            VerticalLine.Layer = "0"

            'add vertical line to CrosshairCollection
            CrosshairCollection.Add(VerticalLine)


            'define points for Horizontal line
            Dim HorizontalLinePoints As New Point2dCollection
            HorizontalLinePoints.Add(New Point2d(0 - CrossLength, 0))
            HorizontalLinePoints.Add(New Point2d(0 + CrossLength, 0))

            'define Horizontal line
            Dim HorizontalLine As New Polyline
            HorizontalLine.AddVertexAt(0, HorizontalLinePoints(0), 0, 0, 0)
            HorizontalLine.AddVertexAt(1, HorizontalLinePoints(1), 0, 0, 0)
            HorizontalLine.Layer = "0"

            'add Horizontal line to CrosshairCollection
            CrosshairCollection.Add(HorizontalLine)


            Return CrosshairCollection   'return the crosshair
        End Function

        Function CreateBlock(ByVal ents As DBObjectCollection, ByVal BlockName As String, ByVal BasePoint As Point3d) As ObjectId

            Dim blkDefs As New ObjectIdCollection()

            Dim btrId As ObjectId

            Dim myDWG As Document = Application.DocumentManager.MdiActiveDocument
            Dim myDB As Database = myDWG.Database
            Dim myED As Editor = myDWG.Editor
            Dim myTrans As Transaction = myDB.TransactionManager.StartTransaction()
            Using myTrans
                ' Get the block table from the drawing
                Dim bt As BlockTable = DirectCast(myTrans.GetObject(myDB.BlockTableId, OpenMode.ForRead), BlockTable)

                ' Create new block table record
                Dim btr As New BlockTableRecord()

                Dim blkName As String = ""
                Do
                    Try
                        ' Validate the provided symbol table name
                        SymbolUtilityServices.ValidateSymbolName(BlockName, False)
                        ' Only set the block name if it isn't in use
                        If bt.Has(BlockName) Then
                            Return bt(BlockName)
                            'myED.WriteMessage(vbLf & "A block with this name already exists.")
                        Else
                            blkName = BlockName
                        End If
                    Catch
                        myED.WriteMessage(Err.Description)
                    End Try
                Loop While blkName = ""


                ' ... and set its properties
                btr.Name = blkName
                ' Add the new block to the block table
                bt.UpgradeOpen()
                btrId = bt.Add(btr)
                myTrans.AddNewlyCreatedDBObject(btr, True)
                ' Add some lines to the block to form a square (the entities belong directly to the block)

                For Each ent As Entity In ents
                    btr.AppendEntity(ent)
                    myTrans.AddNewlyCreatedDBObject(ent, True)
                Next


                ' Commit the transaction
                myTrans.Commit()
            End Using
            Return btrId
        End Function

        Shared Function DrawGridLabels(ByRef myGridSettings As GridSettings, Optional ByVal IsJig As Boolean = False) As DBObjectCollection
            Dim LabelCollection As New DBObjectCollection
            Dim LabelFormat As String = "0.000"
            Dim ExtentsPoint As Point3d = myGridSettings.GridExtents
            Dim GridWidth As Double = myGridSettings.GridExtents.X - myGridSettings.BasePoint.X
            Dim GridHeight As Double = myGridSettings.GridExtents.Y - myGridSettings.BasePoint.Y
            Dim GridInterval As Double = myGridSettings.GridInterval
            Dim BasePoint As Point3d = myGridSettings.BasePoint
            Dim CrossHairSize As Double = myGridSettings.CrosshairSize
            Dim InsertionPoint As Point3d = BasePoint

            Dim NumberOfX As Integer
            NumberOfX = Math.Round(GridWidth / myGridSettings.GridInterval)

            Dim NumberOfY As Integer
            NumberOfY = Math.Round(GridHeight / myGridSettings.GridInterval)

            Dim Label As New MText
            Dim LabelString As String
            Dim LabelTextHeight As Double = 2.5

            Dim LabelOffset As Double = (CrossHairSize / 2) * 1.5

            'Insert right labels
            InsertionPoint = BasePoint
            For Y = 0 To NumberOfY
                'Update insertion point
                InsertionPoint = New Point3d(BasePoint.X + (NumberOfX * GridInterval) + LabelOffset, (BasePoint.Y + (Y * GridInterval)), ExtentsPoint.Y)
                Label = New MText
                If myGridSettings.LocalGrid Then
                    LabelString = (InsertionPoint.Y - myGridSettings.BasePoint.Y).ToString(LabelFormat)
                Else
                    LabelString = InsertionPoint.Y.ToString(LabelFormat)
                End If
                Label.Contents = LabelString
                Label.Location = InsertionPoint
                Label.Rotation = 0
                Label.TextHeight = LabelTextHeight
                Label.Attachment = AttachmentPoint.MiddleLeft
                LabelCollection.Add(Label)
            Next

            'left labels
            InsertionPoint = BasePoint
            For Y = 0 To NumberOfY
                'Update insertion point
                InsertionPoint = New Point3d(BasePoint.X - LabelOffset, (BasePoint.Y + (Y * GridInterval)), BasePoint.Y)
                Label = New MText
                If myGridSettings.LocalGrid Then
                    LabelString = (InsertionPoint.Y - myGridSettings.BasePoint.Y).ToString(LabelFormat)
                Else
                    LabelString = InsertionPoint.Y.ToString(LabelFormat)
                End If
                Label.Contents = LabelString
                Label.Location = InsertionPoint
                Label.Rotation = 0
                Label.TextHeight = LabelTextHeight
                Label.Attachment = AttachmentPoint.MiddleRight
                LabelCollection.Add(Label)
            Next

            'bottom labels
            InsertionPoint = BasePoint
            For X = 0 To NumberOfX
                InsertionPoint = New Point3d(BasePoint.X + (X * GridInterval), BasePoint.Y - LabelOffset, BasePoint.Z)
                Label = New MText
                Label.Location = InsertionPoint
                If myGridSettings.LocalGrid Then
                    LabelString = (InsertionPoint.X - myGridSettings.BasePoint.X).ToString(LabelFormat)
                Else
                    LabelString = InsertionPoint.X.ToString(LabelFormat)
                End If
                Label.Contents = LabelString
                Label.Rotation = Math.PI / 2 '90 degrees in rads
                Label.TextHeight = LabelTextHeight
                Label.Attachment = AttachmentPoint.MiddleRight
                LabelCollection.Add(Label)
            Next

            InsertionPoint = BasePoint
            'top labels
            For X = 0 To NumberOfX
                InsertionPoint = New Point3d(BasePoint.X + (X * GridInterval), BasePoint.Y + (NumberOfY * GridInterval) + LabelOffset, BasePoint.Z)
                Label = New MText
                Label.Location = InsertionPoint
                If myGridSettings.LocalGrid Then
                    LabelString = (InsertionPoint.X - myGridSettings.BasePoint.X).ToString(LabelFormat)
                Else
                    LabelString = InsertionPoint.X.ToString(LabelFormat)
                End If
                Label.Contents = LabelString
                Label.Rotation = Math.PI / 2 '90 degrees in rads
                Label.TextHeight = LabelTextHeight
                Label.Attachment = AttachmentPoint.MiddleLeft
                LabelCollection.Add(Label)
            Next

            Return LabelCollection
        End Function

        Shared Function DrawJigCrossHairs(ByRef myGridSettings As GridSettings, Optional ByVal IsJig As Boolean = False) As DBObjectCollection
            Dim CrosshairCollection As New DBObjectCollection

            Dim ExtentsPoint As Point3d = myGridSettings.GridExtents
            Dim GridWidth As Double = myGridSettings.GridExtents.X - myGridSettings.BasePoint.X
            Dim GridHeight As Double = myGridSettings.GridExtents.Y - myGridSettings.BasePoint.Y
            Dim GridInterval As Double = myGridSettings.GridInterval
            Dim CrossHairBtrId As ObjectId = myGridSettings.BlockID
            Dim BasePoint As Point3d = myGridSettings.BasePoint
            Dim CrossHairSize As Double = myGridSettings.CrosshairSize

            Dim InsertionPoint As Point3d = BasePoint

            Dim NumberOfX As Integer
            NumberOfX = Math.Round(GridWidth / myGridSettings.GridInterval)

            Dim NumberOfY As Integer
            NumberOfY = Math.Round(GridHeight / myGridSettings.GridInterval)

            For Y = 0 To NumberOfY
                For x = 0 To NumberOfX
                    If Y = 0 Or Y = NumberOfY Then
                        Dim CrossHairBlock As New BlockReference(InsertionPoint, CrossHairBtrId)
                        CrossHairBlock.ScaleFactors = New Scale3d(CrossHairSize, CrossHairSize, CrossHairSize)
                        CrosshairCollection.Add(CrossHairBlock)
                    Else
                        If x = 0 Or x = NumberOfX Or IsJig = False Then
                            Dim CrossHairBlock As New BlockReference(InsertionPoint, CrossHairBtrId)
                            CrossHairBlock.ScaleFactors = New Scale3d(CrossHairSize, CrossHairSize, CrossHairSize)
                            CrosshairCollection.Add(CrossHairBlock)
                        End If
                    End If
                    InsertionPoint = New Point3d(InsertionPoint.X + GridInterval, InsertionPoint.Y, InsertionPoint.Z)

                Next
                InsertionPoint = New Point3d(BasePoint.X, InsertionPoint.Y + GridInterval, BasePoint.Z)
            Next

            Return CrosshairCollection
        End Function

    End Class
End Namespace