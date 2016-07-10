Attribute VB_Name = "Plantillas"
Option Explicit

Public ErrNum, ErrDesc 'Para que compile el proyecto.

Private Sub PlantillaParaDelegado()
On Error GoTo DelegateError
    'Código aquí.
    'Si el delegado regresa valor, es buena idea asignarlo de inmediato, si se tiene un valor default.
DelegateError:
    ErrNum = Err.Number: ErrDesc = Err.Description 'ErrNum y ErrDesc deben de estar definidas a nivel módulo.
End Sub

Private Sub PlantillaParaLlamarADelegado()
    'ErrNum y ErrDesc deben de estar definidas y accesibles en el módulo donde esté el delegado (Si fuera otro)
    ErrNum = 0: Application.Run "NombreDelDelegado" ',Más, parámetros, etc
    If ErrNum <> 0 Then Err.Raise ErrNum, , ErrDesc
End Sub
