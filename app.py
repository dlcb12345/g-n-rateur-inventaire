from flask import Flask, request, send_file
import os
import pythoncom
import win32com.client as win32
from werkzeug.utils import secure_filename
from functools import wraps

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

macro_code = '''
# macro_code.py
macro_code = '''
Sub OrganiserTableauAvecIndexEtCouleursUniques()
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim FilePath As String
    Dim LastRow As Long
    Dim LastCol As Long
    Dim i As Long
    Dim ReferenceCol As Integer
    Dim CommentaireCol As Integer
    Dim ColToKeep As Variant
    Dim UniqueIndexes As Object
    Dim IndexKeys As Variant
    Dim Key As Variant
    Dim PasteRow As Long
    Dim ColorIndex As Integer
    Dim GroupColor As Variant
    Dim ResultSheet As Worksheet
    Dim NewSheetName As String
    Dim SheetCount As Integer

    ' Palette de couleurs étendue pour les groupes
    GroupColor = Array(RGB(255, 230, 204), RGB(204, 255, 229), RGB(230, 230, 255), _
                       RGB(255, 255, 204), RGB(255, 204, 255), RGB(229, 204, 255), _
                       RGB(255, 229, 204), RGB(204, 255, 255), RGB(255, 204, 229))

    ' Ouvrir une boîte de dialogue pour sélectionner un fichier Excel
    FilePath = Application.GetOpenFilename(FileFilter:="Fichiers Excel (*.xls; *.xlsx; *.xlsm), *.xls; *.xlsx; *.xlsm", Title:="Sélectionnez un fichier Excel")
    
    ' Vérifier si un fichier a été sélectionné
    If FilePath = "Faux" Then
        MsgBox "Aucun fichier sélectionné. La macro a été annulée.", vbExclamation
        Exit Sub
    End If
    
    ' Ouvrir le fichier sélectionné
    Set wb = Workbooks.Open(FilePath)
    Set ws = wb.Sheets(1) ' Utiliser la première feuille du fichier
    
    ' Trouver la dernière ligne et colonne utilisées
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    LastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Vérifier si la colonne "Reference" existe
    ReferenceCol = 0
    For i = 1 To LastCol
        If LCase(ws.Cells(1, i).Value) = "reference" Then
            ReferenceCol = i
            Exit For
        End If
    Next i

    ' Si la colonne "Reference" n'existe pas, afficher un message et arrêter la macro
    If ReferenceCol = 0 Then
        MsgBox "La colonne 'Reference' est absente. Veuillez vérifier le fichier source.", vbCritical
        Exit Sub
    End If
    
    ' Vérifier si la colonne "Commentaire" existe
    CommentaireCol = 0
    For i = 1 To LastCol
        If LCase(ws.Cells(1, i).Value) = "commentaire" Then
            CommentaireCol = i
            Exit For
        End If
    Next i
    
    ' Ajouter la colonne "Commentaire" en dernière position si elle n'existe pas
    If CommentaireCol = 0 Then
        ws.Cells(1, LastCol + 1).Value = "Commentaire"
        CommentaireCol = LastCol + 1
    End If
    
    ' Définir les colonnes à conserver
    ColToKeep = Array("Reference", "Index utilisateur", "Titre", "N° version", "Date d'édition", _
                      "Date de version", "Date d'application", "Quantité papier", "Commentaire")
    
    ' Supprimer les colonnes non essentielles
    For i = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column To 1 Step -1
        If IsError(Application.Match(ws.Cells(1, i).Value, ColToKeep, 0)) Then
            ws.Columns(i).Delete
        End If
    Next i
    
    ' Mettre à jour la dernière colonne après suppression des colonnes inutiles
    LastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Collecter les Index utilisateurs uniques (extraction avant le tiret)
    Set UniqueIndexes = CreateObject("Scripting.Dictionary")
    For i = 2 To LastRow
        If Not IsEmpty(ws.Cells(i, 2).Value) Then
            ' Extraire uniquement la partie avant le tiret
            Dim IndexUtilisateur As String
            IndexUtilisateur = Split(ws.Cells(i, 2).Value, "-")(0)
            If Not UniqueIndexes.exists(IndexUtilisateur) Then
                UniqueIndexes.Add IndexUtilisateur, IndexUtilisateur
            End If
        End If
    Next i

    ' Trier les Index utilisateurs par ordre alphabétique
    IndexKeys = UniqueIndexes.Keys
    Call BubbleSort(IndexKeys)

    ' Nommer dynamiquement la feuille des résultats
    SheetCount = wb.Sheets.Count
    NewSheetName = "ResultatMacro" & SheetCount + 1
    Set ResultSheet = wb.Sheets.Add
    ResultSheet.Name = NewSheetName

    ' Copier l'en-tête dans la nouvelle feuille
    PasteRow = 1
    ws.Rows(1).Copy Destination:=ResultSheet.Rows(PasteRow)
    PasteRow = PasteRow + 1

    ' Appliquer une couleur différente à chaque groupe d'Index utilisateur
    ColorIndex = 0
    For Each Key In IndexKeys
        ' Ajouter une ligne avec uniquement l'Index utilisateur
        With ResultSheet.Range(ResultSheet.Cells(PasteRow, 1), ResultSheet.Cells(PasteRow, LastCol))
            .Merge
            .Value = Key ' Affiche uniquement l'Index utilisateur (avant le tiret)
            .Font.Bold = True
            .Font.Size = 12
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = GroupColor(ColorIndex Mod (UBound(GroupColor) + 1))
        End With
        PasteRow = PasteRow + 1
        
        ' Copier les données correspondant à l'Index utilisateur
        For i = 2 To LastRow
            Dim CurrentIndex As String
            CurrentIndex = Split(ws.Cells(i, 2).Value, "-")(0) ' Extraire avant le tiret
            If CurrentIndex = Key Then
                ' Colonne Reference
                ResultSheet.Cells(PasteRow, 1).Value = ws.Cells(i, ReferenceCol).Value
                ResultSheet.Cells(PasteRow, 1).Font.Bold = True
                
                ' Copier les autres colonnes demandées
                ResultSheet.Cells(PasteRow, 2).Value = CurrentIndex ' Index utilisateur
                ResultSheet.Cells(PasteRow, 3).Value = ws.Cells(i, 3).Value ' Titre
                ResultSheet.Cells(PasteRow, 4).Value = ws.Cells(i, 4).Value ' N° version
                ResultSheet.Cells(PasteRow, 5).Value = ws.Cells(i, 5).Value ' Date d'édition
                ResultSheet.Cells(PasteRow, 6).Value = ws.Cells(i, 6).Value ' Date de version
                ResultSheet.Cells(PasteRow, 7).Value = ws.Cells(i, 7).Value ' Date d'application
                ResultSheet.Cells(PasteRow, 8).Value = ws.Cells(i, 8).Value ' Quantité papier
                ResultSheet.Cells(PasteRow, 9).Value = ws.Cells(i, CommentaireCol).Value ' Commentaire (dernière colonne)
                PasteRow = PasteRow + 1
            End If
        Next i
        ' Passer à la couleur suivante
        ColorIndex = ColorIndex + 1
    Next Key

    ' Ajuster la largeur des colonnes
    ResultSheet.Columns(1).ColumnWidth = 20 ' Colonne Reference
    ResultSheet.Columns(2).ColumnWidth = 20 ' Colonne Index utilisateur

    ' Ajuster la largeur des colonnes Titre et Commentaire à 660 pixels (50 unités Excel)
    For i = 1 To LastCol
        If LCase(ResultSheet.Cells(1, i).Value) = "titre" Or LCase(ResultSheet.Cells(1, i).Value) = "commentaire" Then
            ResultSheet.Columns(i).ColumnWidth = 50 ' 660 pixels
        End If
    Next i

    ' Appliquer une couleur crème uniquement à la ligne d'en-tête (A1:I1)
    ResultSheet.Range("A1:I1").Interior.Color = RGB(255, 253, 208)

    ' Mise en forme finale
    With ResultSheet.Range("A1:I" & PasteRow - 1)
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Font.Name = "Times New Roman"
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    ' Informer l'utilisateur
    MsgBox "Traitement terminé. Les données ont été exportées avec les filtres et les couleurs demandées.", vbInformation
End Sub

' Fonction pour trier un tableau (méthode Bubble Sort)
Sub BubbleSort(arr As Variant)
    Dim i As Long, j As Long
    Dim temp As Variant
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                temp = arr(j)
                arr(j) = arr(i)
                arr(i) = temp
            End If
        Next j
    Next i
End Sub
'''

'''

def com_thread_init(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        pythoncom.CoInitialize()
        try:
            return func(*args, **kwargs)
        finally:
            pythoncom.CoUninitialize()
    return wrapper

@app.route('/')
def index():
    return '''
    <h1>Upload fichier Excel</h1>
    <form action="/upload" method="post" enctype="multipart/form-data">
      <input type="file" name="file" />
      <input type="submit" value="Convertir et insérer la macro" />
    </form>
    '''

@app.route('/upload', methods=['POST'])
@com_thread_init
def upload_file():
    if 'file' not in request.files:
        return "Pas de fichier envoyé", 400

    file = request.files['file']
    if file.filename == '':
        return "Nom de fichier vide", 400

    filename = secure_filename(file.filename)
    input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(input_path)

    output_filename = os.path.splitext(filename)[0] + '_avec_macro.xlsm'
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)

    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(os.path.abspath(input_path))

        vb_module = wb.VBProject.VBComponents.Add(1)
        vb_module.CodeModule.AddFromString(macro_code)

        xlOpenXMLWorkbookMacroEnabled = 52
        wb.SaveAs(os.path.abspath(output_path), FileFormat=xlOpenXMLWorkbookMacroEnabled)

        wb.Close(SaveChanges=False)
        excel.Quit()

    except Exception as e:
        return f"Erreur Excel : {e}", 500

    return send_file(output_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
