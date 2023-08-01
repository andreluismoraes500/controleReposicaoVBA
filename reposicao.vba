Sub PosiçãoNomeProfessor()
'
' Desenvolvimento Sistema de Reposição de Aulas
'
'
    Dim interador As Integer
    Dim inputProfessor As String
    Dim inputDataReposicao As String
    Dim inputQuantidadeDeAulas As Integer
    
    Dim quantidaDeAulas As Integer
    Dim verificaDomingo As String
    Dim verificaMsgBox As Integer
    
    'ArrayDeStrings (2) tenho que fazer quando existir
    'tenho que verificar se está na coluna H em diante
    
    inputProfessor = UCase(Application.InputBox("Digite o nome do Professor"))
    If inputProfessor = "Falso" Then
        Exit Sub
    End If
    
    inputDataReposicao = Application.InputBox("Digite a data da Reposição")
    If inputDataReposicao = "Falso" Then
        Exit Sub
    End If
    
    interador = 1
    quantidadeDeAulas = 8 'quantidade total de aulas
    verificaDomingo = Weekday(inputDataReposicao) 'se for domingo retorna 1
    
    If verificaDomingo = 1 Then
        MsgBox "Não pode ter aulas aos domingos"
        Exit Sub
    End If
    
    
    Do
        interador = interador + 1
        
        If inputDataReposicao = Range("B" + CStr(interador)).Value Then 'comparando as datas
           'MsgBox "Estou na mesma data"
            If inputProfessor = Range("A" + CStr(interador)).Value Then 'comparando o professor na data acima
                'MsgBox "bateu professor a data com o professor")
                
                If quantidadeDeAulas > Range("C" + CStr(interador)).Value Then
                    MsgBox "posso dar " & quantidadeDeAulas - CInt(Range("C" + CStr(interador)).Value) & " mais aulas neste dia"
                    '6 = apartei sim 7 = apertei não
                    verificaMsgBox = MsgBox("Deseja inserir aula neste dia? - " & CStr(Range("B" + CStr(interador)).Value), vbYesNo, "Inserção de aula")
                    
                    If verificaMsgBox = 6 Then
                        Do
                            'inputQuantidadeDeAulas = Application.InputBox("Digite a quantidade de aulas") substituindo para a função inserindoUmProfessorNaData
                            inserindoUmProfessorNaData interador
                            
                            If CInt(Range("C" + CStr(interador)).Value) + inputQuantidadeDeAulas > 8 Then
                                MsgBox "Você só pode inserir até " & quantidadeDeAulas - CInt(Range("C" + CStr(interador)).Value) & " aulas"
                            Else
                                Range("C" + CStr(interador)) = CInt(Range("C" + CStr(interador)).Value) + inputQuantidadeDeAulas
                                Exit Do
                            End If
                            
                        Loop While CInt(Range("C" + CStr(interador)).Value) + inputQuantidadeDeAulas > 8
                        
                        Exit Do
                    End If
                    
                    Exit Do
                End If
                
                If Range("C" + CStr(interador)).Value >= quantidadeDeAulas Then
                    MsgBox "Não pode colocar aula neste dia"
                    Exit Do
                End If
                
            End If

        End If
    Loop While Range("B" + CStr(interador)).Value <> ""
    
    'inserindo uma nova reposição
    'WTF que linguagem loka kd os parenteses meu paiiiiiiiiii xD
    'para criar a função só realizar o que está abaixo
    insercaoNovoProfessor inputQuantidadeDeAulas, interador, inputProfessor, inputDataReposicao

End Sub

Function insercaoNovoProfessor(inputQuantidadeDeAulas As Integer, interador As Integer, inputProfessor As String, inputDataReposicao As String)
'inserindo uma nova reposição
        Dim inputInstituicao As String
        Dim inputInicioDaAula As String
        Dim inputFimDaAula As String
        Dim quantidadeDeAulas As Integer
        
        If Range("A" + CStr(interador)).Value = "" Then
            Do
                inputInicioDaAula = Application.InputBox("Digite o horario do inicio da aula")
                
                inputFimDaAula = Application.InputBox("Digite o horario do fim da aula")
                                
                If CInt(SubtracaoHoras(CDate(inputInicioDaAula), CDate(inputFimDaAula))) > 8 Then
                    MsgBox "Você só pode inserir até 8 aulas"
                ElseIf CInt(SubtracaoHoras(CDate(inputInicioDaAula), CDate(inputFimDaAula))) = 0 Then
                    Exit Function
                Else

                    quantidadeDeAulas = CInt(SubtracaoHoras(CDate(inputInicioDaAula), CDate(inputFimDaAula)))
                    inputInstituicao = Application.InputBox("Digite a Instituição")
                    Range("A" + CStr(interador)) = UCase(inputProfessor)
                    Range("B" + CStr(interador)) = inputDataReposicao
                    Range("C" + CStr(interador)) = quantidadeDeAulas
                    Range("D" + CStr(interador)) = UCase(inputInstituicao)
                    Range("E" + CStr(interador)) = inputInicioDaAula
                    Range("F" + CStr(interador)) = inputFimDaAula
                End If
            Loop While CInt(SubtracaoHoras(CDate(inputInicioDaAula), CDate(inputFimDaAula))) > 8
        End If
End Function

Function inserindoUmProfessorNaData(interador As Integer)
    'MsgBox "estou aqui inserindo um professor já existente na data"
        Dim inputInstituicao As String
        Dim inputInicioDaAula As String
        Dim inputFimDaAula As String
        Dim quantidadeDeAulas As Integer
        Dim podeInserir As Boolean
        Dim horarioE, horarioF, horarioH, horarioI, horarioK, horarioL As Date
        
        podeInserir = False
        
        quantidadeDeAulas = Range("C" + CStr(interador)).Value
        
        
        horarioE = CDate(Range("E" + CStr(interador)).Value) 'pegando o horario inicio se condição ok passa para o H
        horarioF = CDate(Range("F" + CStr(interador)).Value) 'pegando o horario fim se condição ok passa para o I
        
        horarioH = CDate(Range("H" + CStr(interador)).Value) 'pegando o horario inicio se condição ok passa para o K
        horarioI = CDate(Range("I" + CStr(interador)).Value) 'pegando o horario fim se condição ok passa para o L
        
        horarioK = CDate(Range("K" + CStr(interador)).Value) 'pegando o horario fim se condição ok passa para o N
        horarioL = CDate(Range("L" + CStr(interador)).Value) 'pegando o horario fim se condição ok passa para o O
        
        '------------------------------------------------------------------------- comparando a letra G
        If Range("G" + CStr(interador)).Value = "" Then
            Do
                Do
                    inputInicioDaAula = CDate(Application.InputBox("Digite o horario do inicio da aula"))
                    
                    inputFimDaAula = CDate(Application.InputBox("Digite o horario do fim da aula"))
                    
                    If (inputFimDaAula >= horarioE And horarioF >= inputInicioDaAula) Then 'compara os horario de entrada e saida com horarios já lançados
                        MsgBox "Não pode inserir o professor neste Horario"
                        Else
                        podeInserir = True
                    End If

                Loop While podeInserir = False
                                
                If quantidadeDeAulas + CInt(SubtracaoHoras(CDate(inputInicioDaAula), CDate(inputFimDaAula))) > 8 Then
                    MsgBox "Você só pode inserir até 8 aulas"
                    podeInserir = False
                ElseIf CInt(SubtracaoHoras(CDate(inputInicioDaAula), CDate(inputFimDaAula))) = 0 Then
                    Exit Function
                Else
                    inputInstituicao = Application.InputBox("Digite a Instituição")
                
                    Range("C" + CStr(interador)) = quantidadeDeAulas + CInt(SubtracaoHoras(CDate(inputInicioDaAula), CDate(inputFimDaAula)))
                    Range("G" + CStr(interador)) = UCase(inputInstituicao)
                    Range("H" + CStr(interador)) = inputInicioDaAula
                    Range("I" + CStr(interador)) = inputFimDaAula
                End If
            Loop While quantidadeDeAulas + CInt(SubtracaoHoras(CDate(inputInicioDaAula), CDate(inputFimDaAula))) > 8
            Exit Function
        End If
        
        '------------------------------------------------------------------------- comparando a letra G e J
        If Range("G" + CStr(interador)).Value <> "" And Range("J" + CStr(interador)).Value = "" Then
            'MsgBox "estou aqui agora testando esta nova condição"
            Do
                Do
                    inputInicioDaAula = CDate(Application.InputBox("Digite o horario do inicio da aula"))
                    
                    inputFimDaAula = CDate(Application.InputBox("Digite o horario do fim da aula"))
                    
                    If (inputFimDaAula >= horarioH And horarioI >= inputInicioDaAula) Then
                        MsgBox "Não pode inserir o professor neste Horario"
                        ElseIf (inputFimDaAula >= horarioE And horarioF >= inputInicioDaAula) Then
                            MsgBox "Não pode inserir o professor neste Horario"
                    Else
                        podeInserir = True
                    End If

                Loop While podeInserir = False
                                
                If quantidadeDeAulas + CInt(SubtracaoHoras(CDate(inputInicioDaAula), CDate(inputFimDaAula))) > 8 Then
                    MsgBox "Você só pode inserir até 8 aulas"
                    podeInserir = False
                ElseIf CInt(SubtracaoHoras(CDate(inputInicioDaAula), CDate(inputFimDaAula))) = 0 Then
                    Exit Function
                Else
                    inputInstituicao = Application.InputBox("Digite a Instituição")
                
                    Range("C" + CStr(interador)) = quantidadeDeAulas + CInt(SubtracaoHoras(CDate(inputInicioDaAula), CDate(inputFimDaAula)))
                    Range("J" + CStr(interador)) = UCase(inputInstituicao)
                    Range("K" + CStr(interador)) = inputInicioDaAula
                    Range("L" + CStr(interador)) = inputFimDaAula
                End If
            Loop While quantidadeDeAulas + CInt(SubtracaoHoras(CDate(inputInicioDaAula), CDate(inputFimDaAula))) > 8
            Exit Function
        End If
        
        '------------------------------------------------------------------------- comparando a letra G J e M vou verificar se irá fazer
        
        
End Function


Function SubtracaoHoras(horaInicial As Date, horaFinal As Date) As Double
    Dim horasResultado As Double
    
    'a partir de baixo está funcionando
    
    horasResultado = Round((horaFinal - horaInicial) * 24 * 60, 0) ' Resultado em minutos
    SubtracaoHoras = horasResultado / 50
    
    'MsgBox "A diferença de horas é: " & horasResultado & " minutos"
End Function

'Function ArrayDeStrings(interador As Integer)
    'Declaração do array de strings
    'Dim meusHorarios() As String
'    Dim calculo As Double
    
    'primeiro horario do dia
'    calculo = SubtracaoHoras(Range("E" + CStr(interador)), Range("F" + CStr(interador)))
    
    'segundo horario do dia
'    calculo = calculo + SubtracaoHoras(Range("H" + CStr(interador)), Range("I" + CStr(interador)))
    
    'terceiro horario do dia
'    calculo = calculo + SubtracaoHoras(Range("K" + CStr(interador)), Range("L" + CStr(interador)))
    
    'quarto horario do dia
'    calculo = calculo + SubtracaoHoras(Range("N" + CStr(interador)), Range("O" + CStr(interador)))
'    MsgBox calculo / 50 & " aulas foram dadas neste periodo"
    
'End Function

' Inicializando o array com valores
'meusHorarios = Split("Maçã,Banana,Laranja,Morango", ",")
    
' Acessando elementos individuais do array
'MsgBox "O primeiro elemento do array é: " & meuArray(0)
'MsgBox "O terceiro elemento do array é: " & meuArray(2)
    
' Alterando elementos do array
'meuArray(1) = "Uva"
    
' Percorrendo o array com um loop For
'Dim i As Integer
'For i = LBound(meusHorarios) To UBound(meusHorarios)
'    MsgBox "Elemento " & i & ": " & meusHorarios(i)
'Next i



