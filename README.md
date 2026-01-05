Sub CadUsua()

With Sheets("BASE-USUARIO")
    If Application.WorksheetFunction.CountA(.Range("C3:C1048576")) = 0 Then
    proxID = 1
Else
    proxID = WorksheetFunction.Max(.Range("C3:C1048576")) + 1
End If

.Range("B3:R3").Insert Shift:=xlDown

.Range("C3").Value = proxID


    .Range("D3").Value = Sheets("CADASTRO-CLIENTE").Range("C10").Value
    
    .Range("E3").Value = Split(.Range("D3").Value, " ")(0)
    
    .Range("F3").Value = Sheets("CADASTRO-CLIENTE").Range("C14").Value
    .Range("G3").Value = Sheets("CADASTRO-CLIENTE").Range("C17").Value
    .Range("H3").Value = Sheets("CADASTRO-CLIENTE").Range("G17").Value
    .Range("I3").Value = Sheets("CADASTRO-CLIENTE").Range("C21").Value
    .Range("J3").Value = Sheets("CADASTRO-CLIENTE").Range("C24").Value
    .Range("K3").Value = Sheets("CADASTRO-CLIENTE").Range("G24").Value
    .Range("L3").Value = Sheets("CADASTRO-CLIENTE").Range("C26").Value
    .Range("M3").Value = Sheets("CADASTRO-CLIENTE").Range("G26").Value
    .Range("N3").Value = Sheets("CADASTRO-CLIENTE").Range("C29").Value
    .Range("O3").Value = Sheets("CADASTRO-CLIENTE").Range("C32").Value
    .Range("P3").Value = Sheets("CADASTRO-CLIENTE").Range("C37").Value
    .Range("Q3").Value = Sheets("CADASTRO-CLIENTE").Range("C40").Value
    
    .Range("B3").Value = Now
    .Range("B3").NumberFormat = "dd/mm/yyyy hh:mm:ss"
    
End With
End Sub





Sub cadEmpresa()

With Sheets("BASE-EMPRESA")
    If Application.WorksheetFunction.CountA(.Range("C3:C1048576")) = 0 Then
        proxID = 1
    Else
        proxID = WorksheetFunction.Max(.Range("C3:C1048576")) + 1
    End If
End With

' Para cada opção, se marcada, insere linha e grava dados
If Sheets("CADASTRO-EMPRESA").Range("C38").Value = True Then
    Call gravaEmpresa("BIKE", proxID)
End If

If Sheets("CADASTRO-EMPRESA").Range("D38").Value = True Then
    Call gravaEmpresa("MOTO", proxID)
End If

If Sheets("CADASTRO-EMPRESA").Range("E38").Value = True Then
    Call gravaEmpresa("CARRO", proxID)
End If

If Sheets("CADASTRO-EMPRESA").Range("F38").Value = True Then
    Call gravaEmpresa("VAN", proxID)
End If

End Sub

' --- rotina auxiliar para gravar dados ---
Sub gravaEmpresa(opcao, proxID)

With Sheets("BASE-EMPRESA")
    .Rows(3).Insert Shift:=xlDown

    .Range("B3").Value = Now
    .Range("B3").NumberFormat = "dd/mm/yyyy hh:mm:ss"

    .Range("C3").Value = proxID
    .Range("D3").Value = Sheets("CADASTRO-EMPRESA").Range("C10").Value
    .Range("E3").Value = Split(.Range("D3").Value, " ")(0)
    .Range("F3").Value = Sheets("CADASTRO-EMPRESA").Range("C14").Value
    .Range("G3").Value = Sheets("CADASTRO-EMPRESA").Range("C17").Value
    .Range("H3").Value = Sheets("CADASTRO-EMPRESA").Range("C19").Value
    .Range("I3").Value = Sheets("CADASTRO-EMPRESA").Range("C23").Value
    .Range("J3").Value = Sheets("CADASTRO-EMPRESA").Range("C26").Value
    .Range("K3").Value = Sheets("CADASTRO-EMPRESA").Range("G26").Value
    .Range("L3").Value = Sheets("CADASTRO-EMPRESA").Range("C29").Value
    .Range("M3").Value = Sheets("CADASTRO-EMPRESA").Range("G29").Value
    .Range("N3").Value = Sheets("CADASTRO-EMPRESA").Range("C31").Value
    .Range("O3").Value = Sheets("CADASTRO-EMPRESA").Range("C34").Value

    ' Aqui só uma opção por linha
    .Range("P3").Value = opcao

    .Range("Q3").Value = Sheets("CADASTRO-EMPRESA").Range("C42").Value
    .Range("R3").Value = Sheets("CADASTRO-EMPRESA").Range("C45").Value
End With

End Sub





Sub abreCad()
    Planilha7.Activate
End Sub

Sub abreCadEmpresa()
    Planilha12.Activate
End Sub

Sub Saircont()
    Sheets("MERCADO-LOGADO").Range("T1").MergeArea.ClearContents
    Sheets("MERCADO-LOGADO-JEE").Range("T1").MergeArea.ClearContents
    Sheets("CARRINHO-LOGADO").Range("T1").MergeArea.ClearContents
    Sheets("MERCADO-INICIAL").Activate
End Sub

Sub LogUsu()
    formUsuario2.Show
End Sub

Sub LogEmp()
    formEmpresa.Show
End Sub

Sub IrCarrinhoLog()

    escoTransporte
    Sheets("CARRINHO-LOGADO").Activate
    
    
End Sub

Sub IrCarrinhoSem()

    escoTransporte
    Sheets("CARRINHO-SEM").Activate
End Sub



Sub AdCarrinho()
    produtoID = Split(Application.Caller, "_")(1)
    
    Set cel = Sheets("BASE-PRODUTO").Columns("B").Find( _
        What:=produtoID, LookIn:=xlValues, LookAt:=xlWhole)
        
    If cel Is Nothing Then
        MsgBox "Produto não encontrado"
        Exit Sub
    End If
        
    linhaProduto = cel.Row
    
    Set celCarrinho = Sheets("CARRINHO-LOGADO").Columns("D").Find( _
        What:=produtoID, LookIn:=xlValues, LookAt:=xlWhole)
        
    If Not celCarrinho Is Nothing Then
        celCarrinho.Offset(0, 6).Value = celCarrinho.Offset(0, 6).Value + 1
        celCarrinho.Offset(0, 10).Formula = _
            "=L" & celCarrinho.Row & "*J" & celCarrinho.Row
            
    Else
        linhaCarrinho = Sheets("CARRINHO-LOGADO").Cells(Rows.Count, 4).End(xlUp).Row + 1
    
    Sheets("CARRINHO-LOGADO").Cells(linhaCarrinho, 4).Value = Sheets("BASE-PRODUTO").Cells(linhaProduto, 2).Value
    Sheets("CARRINHO-LOGADO").Cells(linhaCarrinho, 6).Value = Sheets("BASE-PRODUTO").Cells(linhaProduto, 3).Value
    Sheets("CARRINHO-LOGADO").Cells(linhaCarrinho, 8).Value = Sheets("BASE-PRODUTO").Cells(linhaProduto, 8).Value
    Sheets("CARRINHO-LOGADO").Cells(linhaCarrinho, 10).Value = 1
    Sheets("CARRINHO-LOGADO").Cells(linhaCarrinho, 12).Value = Sheets("BASE-PRODUTO").Cells(linhaProduto, 7).Value
    Sheets("CARRINHO-LOGADO").Cells(linhaCarrinho, 14).Formula = "=L" & linhaCarrinho & "*J" & linhaCarrinho
    
    escoTransporte
    

    
    
    End If
    
    MsgBox "Produto adicionado ao carrinho!"
    With Sheets("CARRINHO-LOGADO")
    .Range("U16").Value = Application.Sum(.Range("N5:N" & .Rows.Count))
End With






End Sub


Sub CompraProd()
    nomeBotao = Application.Caller
    
    If InStr(nomeBotao, "_") = 0 Then
        MsgBox "Botao sem ID de produto"
        Exit Sub
    End If
    
    produtoID = Split(nomeBotao, "_")(1)
    
    Set cel = Sheets("BASE-PRODUTO").Columns("B").Find( _
        What:=produtoID, LookIn:=xlValues, LookAt:=xlWhole)
        
    If cel Is Nothing Then
        MsgBox "Produto nao encontrado"
        Exit Sub
    End If
    
    linhaProduto = cel.Row
    
    Set celCarrinho = Sheets("CARRINHO-LOGADO").Columns("D").Find( _
        What:=produtoID, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not celCarrinho Is Nothing Then
        celCarrinho.Offset(0, 6).Value = celCarrinho.Offset(0, 6).Value + 1
        celCarrinho.Offset(0, 10).Formula = _
            "=L" & celCarrinho.Row & "*J" & celCarrinho.Row
    Else
            linhaCarrinho = Sheets("CARRINHO-LOGADO").Cells( _
                Sheets("CARRINHO-LOGADO").Rows.Count, 4).End(xlUp).Row + 1
                
                Sheets("CARRINHO-LOGADO").Cells(linhaCarrinho, 4).Value = _
                Sheets("BASE-PRODUTO").Cells(linhaProduto, 2).Value
                
                Sheets("CARRINHO-LOGADO").Cells(linhaCarrinho, 6).Value = _
                Sheets("BASE-PRODUTO").Cells(linhaProduto, 3).Value
    
                Sheets("CARRINHO-LOGADO").Cells(linhaCarrinho, 8).Value = _
                Sheets("BASE-PRODUTO").Cells(linhaProduto, 8).Value
                
                Sheets("CARRINHO-LOGADO").Cells(linhaCarrinho, 10).Value = _
                1
                
                Sheets("CARRINHO-LOGADO").Cells(linhaCarrinho, 12).Value = _
                Sheets("BASE-PRODUTO").Cells(linhaProduto, 7).Value
                
                Sheets("CARRINHO-LOGADO").Cells(linhaCarrinho, 14).Value = _
                    "=L" & linhaCarrinho & "*J" & linhaCarrinho
                    
                escoTransporte
                
         End If
         
         
         
         
         
         
         
         With Sheets("CARRINHO-LOGADO")
            .Range("U16").Value = Application.Sum(.Range("N5:N" & .Rows.Count))
         End With

         Sheets("CARRINHO-LOGADO").Activate
                
End Sub



Public Sub escoTransporte()

    ' Calcula o peso total (peso unitário x quantidade)
    pesoTotal = Application.WorksheetFunction.SumProduct( _
        Sheets("CARRINHO-LOGADO").Range("H5:H100"), _
        Sheets("CARRINHO-LOGADO").Range("J5:J100"))
    
    riscoAlto = False

    ' Verifica risco dos produtos no carrinho
    For linha = 5 To 100
        
        idCarrinho = Sheets("CARRINHO-LOGADO").Cells(linha, "D").Value
        
        If idCarrinho <> "" Then
            
            On Error Resume Next
            risco = Application.WorksheetFunction.VLookup( _
                        idCarrinho, _
                        Sheets("BASE-PRODUTO").Range("B3:M42"), _
                        12, False)
            On Error GoTo 0
            
            If risco = "Alto" Then
                riscoAlto = True
                Exit For
            End If
            
        End If
        
    Next linha

    ' Define transporte automaticamente
    If riscoAlto = True Then
        
        If pesoTotal > 20 Then
            transporte = "Van"
        Else
            transporte = "Carro"
        End If
    
    ElseIf pesoTotal > 20 Then
        transporte = "Van"
    
    ElseIf pesoTotal > 6 Then
        transporte = "Carro"
    
    ElseIf pesoTotal > 2 Then
        transporte = "Moto"
    
    Else
        transporte = "Bicicleta"
    
    End If

    ' Aplica o resultado na célula
    Sheets("CARRINHO-LOGADO").Range("T12").Value = transporte

End Sub


Public Sub escoTransporteSem()

    ' Calcula o peso total (peso unitário x quantidade)
    pesoTotal = Application.WorksheetFunction.SumProduct( _
        Sheets("CARRINHO-SEM").Range("H5:H100"), _
        Sheets("CARRINHO-SEM").Range("J5:J100"))
    
    riscoAlto = False

    ' Verifica risco dos produtos no carrinho
    For linha = 5 To 100
        
        idCarrinho = Sheets("CARRINHO-SEM").Cells(linha, "D").Value
        
        If idCarrinho <> "" Then
            
            On Error Resume Next
            risco = Application.WorksheetFunction.VLookup( _
                        idCarrinho, _
                        Sheets("BASE-PRODUTO").Range("B3:M42"), _
                        12, False)
            On Error GoTo 0
            
            If risco = "Alto" Then
                riscoAlto = True
                Exit For
            End If
            
        End If
        
    Next linha

    ' Define transporte automaticamente
    If riscoAlto = True Then
        
        If pesoTotal > 20 Then
            transporte = "Van"
        Else
            transporte = "Carro"
        End If
    
    ElseIf pesoTotal > 20 Then
        transporte = "Van"
    
    ElseIf pesoTotal > 6 Then
        transporte = "Carro"
    
    ElseIf pesoTotal > 2 Then
        transporte = "Moto"
    
    Else
        transporte = "Bicicleta"
    
    End If

    ' Aplica o resultado na célula
    Sheets("CARRINHO-SEM").Range("T12").Value = transporte

End Sub



Sub AdCarrinhoSem()
    produtoID = Split(Application.Caller, "_")(1)
    
    Set cel = Sheets("BASE-PRODUTO").Columns("B").Find( _
        What:=produtoID, LookIn:=xlValues, LookAt:=xlWhole)
        
    If cel Is Nothing Then
        MsgBox "Produto não encontrado"
        Exit Sub
    End If
        
    linhaProduto = cel.Row
    
    Set celCarrinho = Sheets("CARRINHO-SEM").Columns("D").Find( _
        What:=produtoID, LookIn:=xlValues, LookAt:=xlWhole)
        
    If Not celCarrinho Is Nothing Then
        celCarrinho.Offset(0, 6).Value = celCarrinho.Offset(0, 6).Value + 1
        celCarrinho.Offset(0, 10).Formula = _
            "=L" & celCarrinho.Row & "*J" & celCarrinho.Row
            
    Else
        linhaCarrinho = Sheets("CARRINHO-SEM").Cells(Rows.Count, 4).End(xlUp).Row + 1
    
    Sheets("CARRINHO-SEM").Cells(linhaCarrinho, 4).Value = Sheets("BASE-PRODUTO").Cells(linhaProduto, 2).Value
    Sheets("CARRINHO-SEM").Cells(linhaCarrinho, 6).Value = Sheets("BASE-PRODUTO").Cells(linhaProduto, 3).Value
    Sheets("CARRINHO-SEM").Cells(linhaCarrinho, 8).Value = Sheets("BASE-PRODUTO").Cells(linhaProduto, 8).Value
    Sheets("CARRINHO-SEM").Cells(linhaCarrinho, 10).Value = 1
    Sheets("CARRINHO-SEM").Cells(linhaCarrinho, 12).Value = Sheets("BASE-PRODUTO").Cells(linhaProduto, 7).Value
    Sheets("CARRINHO-SEM").Cells(linhaCarrinho, 14).Formula = "=L" & linhaCarrinho & "*J" & linhaCarrinho
    
    
    escoTransporteSem

    
    
    End If
    
    MsgBox "Produto adicionado ao carrinho!"
    With Sheets("CARRINHO-SEM")
    .Range("U16").Value = Application.Sum(.Range("N5:N" & .Rows.Count))
End With






End Sub


Sub CompraProdSem()
    nomeBotao = Application.Caller
    
    If InStr(nomeBotao, "_") = 0 Then
        MsgBox "Botao sem ID de produto"
        Exit Sub
    End If
    
    produtoID = Split(nomeBotao, "_")(1)
    
    Set cel = Sheets("BASE-PRODUTO").Columns("B").Find( _
        What:=produtoID, LookIn:=xlValues, LookAt:=xlWhole)
        
    If cel Is Nothing Then
        MsgBox "Produto nao encontrado"
        Exit Sub
    End If
    
    linhaProduto = cel.Row
    
    Set celCarrinho = Sheets("CARRINHO-SEM").Columns("D").Find( _
        What:=produtoID, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not celCarrinho Is Nothing Then
        celCarrinho.Offset(0, 6).Value = celCarrinho.Offset(0, 6).Value + 1
        celCarrinho.Offset(0, 10).Formula = _
            "=L" & celCarrinho.Row & "*J" & celCarrinho.Row
    Else
            linhaCarrinho = Sheets("CARRINHO-SEM").Cells( _
                Sheets("CARRINHO-SEM").Rows.Count, 4).End(xlUp).Row + 1
                
                Sheets("CARRINHO-SEM").Cells(linhaCarrinho, 4).Value = _
                Sheets("BASE-PRODUTO").Cells(linhaProduto, 2).Value
                
                Sheets("CARRINHO-SEM").Cells(linhaCarrinho, 6).Value = _
                Sheets("BASE-PRODUTO").Cells(linhaProduto, 3).Value
    
                Sheets("CARRINHO-SEM").Cells(linhaCarrinho, 8).Value = _
                Sheets("BASE-PRODUTO").Cells(linhaProduto, 8).Value
                
                Sheets("CARRINHO-SEM").Cells(linhaCarrinho, 10).Value = _
                1
                
                Sheets("CARRINHO-SEM").Cells(linhaCarrinho, 12).Value = _
                Sheets("BASE-PRODUTO").Cells(linhaProduto, 7).Value
                
                Sheets("CARRINHO-SEM").Cells(linhaCarrinho, 14).Value = _
                    "=L" & linhaCarrinho & "*J" & linhaCarrinho
                    
                    
                escoTransporteSem
                
         End If
         
         
         
         
         
         
         
         With Sheets("CARRINHO-SEM")
            .Range("U16").Value = Application.Sum(.Range("N5:N" & .Rows.Count))
         End With

         Sheets("CARRINHO-SEM").Activate
                
End Sub

