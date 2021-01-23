'*************************************************************************************************
'Script Description     :Main script which calls different relative modules to complete the flow
'Test Tool/Version		: Unified Functional Testing 14.50
'Application Automated	: Siebel CRM
'Author				    : Manoj Kumar Gangadharan
'Date Created			: 11/12/2018 (mm/dd/yyyy)
'Date last Modified     : 11/22/2018
'**************************************************************************************************
strFileName="..\dtConComplemento\InstallationWowSiebelTOA_ConComplemento_JIEC.xlsx"
If Datatable.GetSheet("Installation").GetCurrentRow = 1 Then

'* Load the Environment Variables includes Siebel URL, Username and Password
    fnLoadEnvironmentVariables()

    '*Creation of sheets
    DataTable.AddSheet("ProfileCreation")
    DataTable.AddSheet("ProductAvailability")
    DataTable.AddSheet("ProductSelection")
    DataTable.AddSheet("Output")
    
    
    '*Import of data
    Datatable.ImportSheet strFileName,"Installation", "Installation"
    Datatable.ImportSheet strFileName,"ProfileCreation", "ProfileCreation"
    Datatable.ImportSheet strFileName,"ProductAvailability", "ProductAvailability"
    Datatable.ImportSheet strFileName,"ProductSelection", "ProductSelection"
    Datatable.ImportSheet strFileName,"Output", "Output"
     
    call fnDisplay("p_Execute", "Installation")
    print  "                                     import      "
'    If i>1 Then
'      PRINT "saLIR "
'      Reporter.ReportEvent micFail, "Siebel Error", "sALIR POR QUE SON MUCHOS INTENTOS"
'    ExitAction
'    End If    
End  If
' ADD20200621 JIEC mostrar los status ejecutados

If Datatable.GetSheet("Installation").GetRowCount= Datatable.GetSheet("Installation").GetCurrentRow Then
	call fnDisplay("p_Execute", "Installation")
End If
'* To check Whether the Script to be exuected or not with the flag value
'...................:::::::OPCIONES     ::::::::::.................................
'  Y--> DE INCIO A FIN LA CUENTA                                   ::::::::::::::::
'  PF--> CREA LA CUENTAS HASTA SELECCIONAR PRODUCTOS               :::::::::::::::: 
'  PE ---> PEGAR EQUIPOS, PROGRAMAR Y ENVIAR                       ::::::::::::::::
'  SEP --> TEST PROBAR SI SELECCIONAR PRODUCTOS(SOLAMENTE DEV)     ::::::::::::::::
'...................................................................................

If DataTable.Value ("p_Execute", "Installation") = uCase("Y") or DataTable.Value ("p_Execute", "Installation") = uCase("PF") or DataTable.Value ("p_Execute", "Installation") = uCase("PE") or DataTable.Value ("p_Execute", "Installation") = uCase("SEP") Then    
'If DataTable.Value ("p_Execute", "Installation") = uCase("SEP") Then    

		 '*Create the data table objects
		    Set dtInstall = Datatable.GetSheet("Installation")
		    Set dtProfileCreation = Datatable.GetSheet("ProfileCreation")
		    Set dtProductAvailability = Datatable.GetSheet("ProductAvailability")
		    Set dtProductSelection = Datatable.GetSheet("ProductSelection")
		    Set dtOutput = Datatable.GetSheet("Output")
		    
		    
		    '*Set the current rows for each Tab based on the current row of the Casos tab
		    dtProfileCreation.SetCurrentRow(dtInstall.GetCurrentRow)
		    dtProductAvailability.SetCurrentRow(dtInstall.GetCurrentRow)
		    dtProductSelection.SetCurrentRow(dtInstall.GetCurrentRow)
		    dtOutput.SetCurrentRow(dtInstall.GetCurrentRow)
		     
		     print dtInstall.GetCurrentRow&"  ......::::::::  "&DataTable.Value ("TestCase", "Installation") 
             strLog=dtInstall.GetCurrentRow&"  ......::::::::  "&DataTable.Value ("TestCase", "Installation") &"  "&DataTable.Value ("p_Execute", "Installation")
             fnUtilWriteLog strLog  
		 
			Systemutil.CloseProcessByName("SiebelAX_Test_Automation_21233.exe")
				sUserName = Environment.Value ("gUserProfileiZZi")
				sPassword = Environment.Value ("gPasswordProfileiZZi")
			Call LoginToSiebel(sUserName, sPassword)
			
			DataTable.Value ("o_TestCase", "Output") =DataTable.Value ("TestCase", "Installation") 
			

				Select Case DataTable.Value ("p_Execute", "Installation") 
				Case "Y"
				        print "                  "&DataTable.Value ("p_Execute", "Installation")
				        Call fnProfilecreationNew()
						Call fnAddressCreateWow()
						Call fnBillingType()
						Call fnProductSelection()		
						Call fnPersonalizerWow("ProductSelection")
						
						'      'ADD20191101 JIEC personalizar se cambia  por todolo de abajo
						       Call fhPersonalizerDatosTOA("ProductSelection")
						       wait 5
						       pagoAdd = DataTable.Value ("p_Pago", "ProfileCreation")
						       'ADD12012021 JLSW si es pago total o parcial se valida esta variable en la data pagoADD
						       If pagoAdd <> "" Then
						       	cuenta = SiebApplication("Siebel Communications").SiebScreen("Ordenes de Servicio").SiebView("Detalles").SiebApplet("Orden de servicio").SiebText("Numero Cuenta").GetROProperty("text")
                                order = SiebApplication("Siebel Communications").SiebScreen("Ordenes de Servicio").SiebView("Detalles").SiebApplet("Orden de servicio").SiebText("Número").GetROProperty("text")
                                pago = SiebApplication("Siebel Communications").SiebScreen("Ordenes de Servicio").SiebView("Detalles").SiebApplet("Orden de servicio").SiebCurrency("Total").GetROProperty("text") 
                               If pagoAdd = "Parcial" Then
                               	pago = Mid(pago,1,(Len(pago))-1)
                               End If
                               Call fnApplyPaymentAnticipadoMonto(cuenta,pago)
                               Call fnBuscarOrden(order)
                               Set oScroll = Createobject("Wscript.shell")
		                       oScroll.SendKeys "{PgDn}"
		                       oScroll.SendKeys "{PgDn}"
		                       Set oScroll = Nothing
		
		                        'Click on Expandir items from Detalles Tab
		                       SiebApplication("Siebel Communications").SiebScreen("Ordenes de Servicio").SiebView("Detalles").SiebApplet("Detalles").SiebButton("Expandir Ítems").Click
		
		                       'Scroll down to get the clear view on the Productlineitems
		                        Set oLineitemScroll = Createobject("Wscript.shell")
		                        oLineitemScroll.SendKeys "{DOWN}"
		                        oLineitemScroll.SendKeys "{DOWN}"
                               wait 5

						       End If
'izzi Telefonia,Hazlo 2 lineas
						Call fnFinalSubmissionWowTOA("ProductSelection") ' Pegar Telefono si tuviera,programar y pegar tag - contrato
								
				
								
				Case "DEGUB"
					print "                  "&DataTable.Value ("p_Execute", "Installation")
					
							   If fnFindOS(DataTable.Value ("o_OrderService", "Output")) Then
									 
									call fnProgramarFechaAtencion()  'Fecha de atencion
 
								End If
								
								
			
				Case "PF"
					print "                  "&DataTable.Value ("p_Execute", "Installation")
								Call fnProfilecreationNew()
								Call fnAddressCreateWow()
								Call fnBillingType()
								Call fnProductSelection()
								Call fnPersonalizerWow("ProductSelection")
								
						'      'ADD20191101 JIEC personalizar se cambia  por todolo de abajo
						       Call fhPersonalizerDatosTOA("ProductSelection")

						
				Case "CP"
					print "                  "&DataTable.Value ("p_Execute", "Installation")
					strMsg="-"
		            Call fnUtilDTWriterMsg("Installation","p_Execute",strMsg)
		        
								Call fnProfilecreationNew()
								Call fnAddressCreateWow()
								Call fnBillingType()
								Call fnProductSelection()
										
					             
					              
					
					
				Case Else
					print "         No existe opcion:          "&DataTable.Value ("p_Execute", "Installation")
					ExitActioniteration	
					
				End Select
		
	    print  "        -----  "& DataTable.Value ("o_Status", "Output")   	
		strMsg=" "&DataTable.Value ("o_Status", "Output")
		Call fnUtilDTWriterMsg("Installation","p_Execute",strMsg)
		 
		 strLog=DataTable.Value ("o_TestCase", "Output")&" "&DataTable.Value ("o_Numeroid", "Output")&" "&DataTable.Value ("o_OrderService", "Output")&" "&DataTable.Value ("o_Status", "Output")&" "&DataTable.Value ("o_ContractoNo", "Output")
         fnUtilWriteLog strLog  
         
		print "                   .."
	    Call fnExportToExcel(strFileName,"Installation", "Installation")
	    Call fnExportToExcel(strFileName,"Output", "Output")
         print "                                   .."		
		
Else
	'If the Flag is not "Y" then exit the Test
	ExitActionIteration

End if
'----------------------------------------------------------------
'----------------------------------------------------------------
'----------------------------------------------------------------




'SiebApplication("Siebel Communications").SiebScreen("Ordenes de Servicio").SiebView("Detalles").SiebApplet("Fecha/Hora de atención").SiebButton("Ok").Click



'SiebApplication("Siebel Communications").SiebScreen("Ordenes de Servicio").SiebView("Detalles").SiebApplet("test1").SiebButton("Ok").Click

'
'wait 5
'  fnLoadEnvironmentVariables()
'cuenta = SiebApplication("Siebel Communications").SiebScreen("Ordenes de Servicio").SiebView("Detalles").SiebApplet("Orden de servicio").SiebText("Numero Cuenta").GetROProperty("text")
'order = SiebApplication("Siebel Communications").SiebScreen("Ordenes de Servicio").SiebView("Detalles").SiebApplet("Orden de servicio").SiebText("Número").GetROProperty("text")
'pago = SiebApplication("Siebel Communications").SiebScreen("Ordenes de Servicio").SiebView("Detalles").SiebApplet("Orden de servicio").SiebCurrency("Total").GetROProperty("text") 
'                               Call fnApplyPaymentAnticipadoMonto(cuenta,pago)
'                               Call fnBuscarOrden(order)
'                               wait 5
'                               


'wait 5
'
'                         StrValue2="Single Play;"
'				 	      Set Testdesc1=description.create
'						 Testdesc1("micClass").value= "WebTable"
'						 Testdesc1("column names").value =StrValue2  
'			
'			             set sChildSPWebTableObj =Browser("Siebel Communications_Numero").Page("Siebel Communications").Frame("CfgMainFrame Frame").ChildObjects(Testdesc1)
'						 print "h3 "&sChildSPWebTableObj.count
'						 sChildSPWebTableObj(0).highlight()
'						 
'						  Set TestdescWT2=description.create
'						 TestdescWT2("micClass").value= "WebTable"
'						 'TestdescWT2("outertext").value =pStrValue&".*"  	
'						 
'						 set sChild2SPWebTableObj =sChildSPWebTableObj(0).ChildObjects(TestdescWT2)
'						 sChild2SPWebTableObj(0).highlight()
'						  print "h2 "&sChild2SPWebTableObj.count
'					     fnSelectWebTableValueContinue sChild2SPWebTableObj(0), "SP PackTV" 'sp_var(i)
