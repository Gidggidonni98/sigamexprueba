
var MENSUAL_ISR ={
    SI(C16="","",(SI( (((((C16)-(BUSCARV(C16,ISR_Mensual, 1,VERDADERO)))*((BUSCARV(C16,ISR_Mensual, 4,VERDADERO))/100))+(BUSCARV(C16,ISR_Mensual, 3,VERDADERO)))-(BUSCARV(C16,Subsidio_Mensual, 3,VERDADERO))) < 0, 0,(((((C16)-(BUSCARV(C16,ISR_Mensual, 1,VERDADERO)))*((BUSCARV(C16,ISR_Mensual, 4,VERDADERO))/100))+(BUSCARV(C16,ISR_Mensual, 3,VERDADERO)))-(BUSCARV(C16,Subsidio_Mensual, 3,VERDADERO))))))
}

Function ISR(sueldo, isrMensual, subsidioM)

    If sueldo = "" Then
    ISR = ""
    Else
If (((((sueldo) - (Application.WorksheetFunction.VLookup(sueldo, isrMensual, 1, True))) * ((Application.WorksheetFunction.VLookup(sueldo, isrMensual, 4, True) / 100))) + (Application.WorksheetFunction.VLookup(sueldo, isrMensual, 3, True))) - (Application.WorksheetFunction.VLookup(sueldo, subsidioM, 3, True))) < 0 Then
                                                                                                                                                                                                                                                               
                                                                
    ISR = 0
    
    Else
    
    ISR = (((((sueldo) - (Application.WorksheetFunction.VLookup(sueldo, isrMensual, 1, True))) * ((Application.WorksheetFunction.VLookup(sueldo, isrMensual, 4, True) / 100))) + (Application.WorksheetFunction.VLookup(sueldo, isrMensual, 3, True))) - (Application.WorksheetFunction.VLookup(sueldo, subsidioM, 3, True)))
    
    
    End If
    
    
    
    End If
    

End Function


I = ((((a+b)*(d/100))+(f))-(h));

i+h = ((a+b)*(d/100))+ (f)

(i+h)/ (f) = (a+b)*(d*100)

((i+h)/(f))/(d*100) = a+b

(((i+h)/f)/(d*100)- (b)) = a


=SI(D17="MENSUAL",ISR(C17,isr_mensual,subsidio_mensual),SI(D17="QUINCENAL",ISR(C17,isr_quincenal,subsidio_qunicenal),SI(D17 = "SEMANAL",ISR(C17,isr_semanal,subsidio_semanal),ISR(C17,isr_catorcenal,subsidio_catorcenal))))

