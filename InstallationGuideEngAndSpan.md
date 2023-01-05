Macro Guide
This is how the macro works, in case it is deleted. That way it is still possible to do work without it.
Copy LM serial numbers of MustGo’s to a new spreadsheet  copy LM serial numbers of Expedites to the same spreadsheet below MustGo’s  double underline and/or bold the Expedites  delete header row  conditional formatting  duplicate values  data  filter  filter by color  manually delete all; this gives you an opportunity to make sure the correct serial numbers have been selected  clear fitler
Así funciona la macro, en caso de que se elimine. De esa manera, todavía es posible trabajar sin él.
Copie los LM serial numbers de MustGo's en una nueva spreadsheet  copie los números de LM serial numbers de Expedites en el mismo spreadsheet debajo de MustGo's  doble subrayado y/o bold los Expedites  eliminar header row  conditional formatting  duplicate values   data  filter  filter by color  eliminar manualmente todo; esto le da la oportunidad de asegurarse de que se han seleccionado los serial numbers correctos  quitar el filtro


Code for macro 1 / Codigo para macro 1 
Sub IDENTIFYDUPLICATE()
	'
	' IDENTIFYDUPLICATE Macro
	'
	
	'
	Columns("A:A").Select
	Selection.FormatConditions.AddUniqueValues
	Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
	Selection.FormatConditions(1).DupeUnique = xlDuplicate
	With Selection.FormatConditions(1).Font
	.Color = -16383844
	.TintAndShade = 0
	End With
	With Selection.FormatConditions(1).Interior
	.PatternColorIndex = xlAutomatic
	.Color = 13551615
	.TintAndShade = 0
	End With
	Selection.FormatConditions(1).StopIfTrue = False
	Selection.AutoFilter
	ActiveSheet.Range("$A$1:$A$2179").AutoFilter Field:=1, Criteria1:=RGB(255, _
	199, 206), Operator:=xlFilterCellColor
	End Sub

 

Link to GitHub for code: https://github.com/robleseli/claudiaexcelmacros/blob/main/DelDuplicateValues.vbs
Github is a page to store code and scripts. There is no need to have an account to access these codes. My account is robleseli
Github es una pagina para almacenar códigos y scripts. No hay necesidad de tener cuenta para acessar estos códigos. Mi cuenta es robleseli


Installing the Macro:

NOTES:
-The “developer tab” must be enabled in order to access, save, and record macros. If you don’t see it, you must enable it.
-if a dialog box is asking about macro-enabled workbooks, select YES. If you select no, then you will not be able to run or access macros 
-If a workbook called PERSONAL.xlsx opens, leave it open. This is a default workbook that Excel uses to store your macros. Do not edit or tamper with it. 
You need a “personal workbook” to store all of your macros, that way you can access them whenever you open any workbook. Do not edit or add anything to this workbook. Do not close it if it opens. 

First, enable the developer tab: 
File  Options  Customize ribbon  Customize the ribbon  Main tabs  check “developer” 

Then, enable personal workbooks:
File   Options  Add-ins  Disabled Items   click Go   Disabled Items dialog box  Personal Macro Workbook  click Enable   Restart your Excel  View   Window   Unhide
When creating a macro, save them to this workbook. 

Create a random macro:
Developer  Code  Record macro  Name the macro  No shortcut is necessary, but you may add one  No description is necessary, but you may add one.  Ok  Perform random action; click on a cell.  Developer  Stop Recording  Macros  Select your macro  Edit  Delete the code from the macro you made  Paste my code into it  Save  Make sure it is saving to PERSONAL.xlsx 

You have now recorded a macro to your excel workbooks. 

Run the macro:
Copy MustGos  Paste Expedites underneath  Run Macro  Delete highlighted values  remove filter  remove leftover expedites  pivot table 



NOTAS:	
-La "developer tab" debe estar habilitada para acceder, guardar y grabar macros. Si no lo ves, debes habilitarlo.
- si aparece uma notificacion que pregunta sobre macro-enabled workbooks, seleccione SÍ. Si selecciona no, entonces no podrá ejecutar ni acceder a los macros
- Si se abre un libro llamado PERSONAL.xlsx, déjelo abierto. Este es un libro de trabajo predeterminado que utiliza Excel para almacenar sus macros. No lo edite ni lo manipule.
Necesita un "personal" para almacenar todos sus macros, de esa manera puede acceder cada vez que abre cualquier workbook. No edite ni agregue nada a este libro de trabajo. No lo cierres si se abre
Primero, habilite la pestaña de desarrollador:
File  Options  Customize ribbon  Customize the ribbon  Main tabs  seleccionar “developer” 

Luego, habilite el personal workbook:
File   Options  Add-ins  Disabled Items   click Go   Disabled Items dialog box  Personal Macro Workbook  click Enable. Re iniciar Excel  View   Window   Unhide.
Ahora ha habilitado el personal workbook. Al crear un macro, guárdela en este workbook.

Crea una macro aleatoria:
Developer  Code  Record macro Nombre de la macro  No es necesario un shortcute, pero puede agregar uno  No es necesaria una descripción, pero puede agregar una.  Ok  Realizar acción aleatoria; haga clic en una cell.  Developer  Dejar de grabar  Elimine el código de la macro que creó   Pegue mi código   Save   Asegúrese de que se guarde en PERSONAL.xlsx

Ahora ha grabado una macro en sus libros de Excel.

Run the macro:
Copy MustGos  Paste Expedites underneath  Run Macro  Delete highlighted values  remove filter  remove leftover expedites  pivot table 

