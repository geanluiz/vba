Attribute VB_Name = "Valores"
Option Explicit

'parametros das funcoes
Dim material As Variant

Dim multi As String

Dim sLarg As Single
Dim sAlt As Single
Dim sProf As Single
Dim qtLatS As Integer
Dim qtRipa As Integer

Dim iLarg As Single
Dim iAlt As Single
Dim iProf As Single
Dim qtLatI As Integer
Dim qtRipaF As Integer

Dim largPS As Single

Dim largPI As Single
Dim moldRpdAplqPI As String

Dim largG As Single
Dim moldRpdAplqG As String
Dim profG As Single

Dim modelo As String
Dim largS As Single
Dim largI As Single
Dim altS As Single
Dim altI As Single
Dim profS As Single
Dim profI As Single
Dim mold As String
Dim qtPInf As Integer


'valor dos materiais
Function vMaterial(material) As Single
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
    Dim vChapas As ListObject: Set vChapas = ws.ListObjects("ValoresChapas")
    Dim vAcess As ListObject: Set vAcess = ws.ListObjects("ValoresAcess")

    Dim m2chapa As Single
    m2chapa = vChapas.DataBodyRange.Cells(1, 2).Value

    Select Case material
        Case 5: vMaterial = vChapas.DataBodyRange.Cells(1, 3).Value / m2chapa
        Case 10: vMaterial = vChapas.DataBodyRange.Cells(1, 4).Value / m2chapa
        Case 15: vMaterial = vChapas.DataBodyRange.Cells(1, 5).Value / m2chapa
        Case 20: vMaterial = vChapas.DataBodyRange.Cells(1, 6).Value / m2chapa
        Case "dobs": vMaterial = vAcess.DataBodyRange.Cells(1, 1).Value
        Case "pux": vMaterial = vAcess.DataBodyRange.Cells(1, 2).Value
        Case "laca": vMaterial = vAcess.DataBodyRange.Cells(1, 3).Value
        Case "esp": vMaterial = vAcess.DataBodyRange.Cells(1, 4).Value
        Case "corr": vMaterial = vAcess.DataBodyRange.Cells(1, 5).Value
        Case "pe":  vMaterial = vAcess.DataBodyRange.Cells(1, 6).Value
    End Select

End Function

'retorna o markup conforme tipo de material
Function markup(multi) As Single

    Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
    Dim dadosCliente As ListObject
    Set dadosCliente = ws.ListObjects("DadosOrcto")
    Dim mkp As Integer

    mkp = dadosCliente.DataBodyRange.Cells(1, 4).Value

    Select Case multi
        Case 3: markup = mkp
        Case 2: markup = mkp * 0.8
    End Select
    
End Function

'valor da caixa superior
Function vCxSup(sLarg, sAlt, sProf, qtLatS, qtRipa) As Single

    Const QTBASE As Integer = 2
    Const QTFUNDO As Integer = 1
    Const QTPRAT As Integer = 2
    Const largRipa As Single = 0.05
    
    Dim sLats As Single
    Dim sBase As Single
    Dim sFundo As Single
    Dim sPrats As Single
    Dim sRipaF As Single
    Dim vFinal05 As Single
    Dim vFinal15 As Single
    Dim vFinal20 As Single
    Dim finalCxSup As Single
    Dim m2Laca As Single
    Dim lacaSup As Single

    
    'calcula m2 da caixa
    sLats = qtLatS * sAlt * sProf
    sBase = QTBASE * sLarg * sProf
    sFundo = QTFUNDO * sLarg * sAlt
    sPrats = QTPRAT * sLarg * sProf
    sRipaF = qtRipa * largRipa * sAlt
    
    'valor chapas
    vFinal05 = sFundo * vMaterial(5)
    vFinal15 = (sLats + sPrats) * vMaterial(15)
    vFinal20 = (sBase + sRipaF) * vMaterial(20)
    
    finalCxSup = (vFinal05 + vFinal15 + vFinal20) * markup(3)
    
          
    'calcula pintura da caixa
    m2Laca = (sLats * 2) + (sBase * 2) + sFundo + (sPrats * 2) + sRipaF
    lacaSup = m2Laca * vMaterial("laca")
    

    vCxSup = Round(finalCxSup + lacaSup, 2)

End Function

'valor da caixa inferior
Function vCxInf(iLarg, iAlt, iProf, qtLatI, qtRipaF) As Single

    Const QTBASE As Integer = 1
    Const QTFUNDO As Integer = 1
    Const QTPRAT As Integer = 1
    Const LARGRIPAF As Single = 0.05
    Const LARGRIPAE As Single = 0.07
    
    Dim iLats As Single
    Dim iBase As Single
    Dim iFundo As Single
    Dim iPrats As Single
    Dim iRipaF As Single
    Dim iRipaE As Single
    Dim iRod As Single
    Dim vFinal05 As Single
    Dim vFinal15 As Single
    Dim vFinal20 As Single
    Dim finalCxInf As Single
    Dim m2Laca As Single
    Dim lacaInf As Single
    Dim iAcess As Single
    Dim qtPes As Integer
    Dim qtRipaProf As Integer

    
    'calcula m2 da caixa
    iLats = qtLatI * iAlt * iProf
    iBase = QTBASE * iLarg * iProf
    iFundo = QTFUNDO * iLarg * iAlt
    iPrats = QTPRAT * iLarg * iProf
    iRipaF = qtRipaF * LARGRIPAF * iAlt

    
    If iLarg > 0.8 Then qtRipaProf = 3 Else qtRipaProf = 2
    iRipaE = ((iAlt * 4) + (iLarg * 3) + (iProf * qtRipaProf)) * LARGRIPAE

    iRod = 2 * (iLarg + iProf) * LARGRIPAE


    'valor chapas
    vFinal05 = iFundo * vMaterial(5)
    vFinal15 = (iLats + iPrats + iRipaE + iRod) * vMaterial(15)
    vFinal20 = (iBase + iRipaF) * vMaterial(20)
    
    finalCxInf = (vFinal05 + vFinal15 + vFinal20) * markup(3)
    
          
    'calcula pintura da caixa
    m2Laca = (iLats * 2) + iBase + iFundo + (iPrats * 2) + iRipaF + iRipaE + iRod
    lacaInf = m2Laca * vMaterial("laca")


    'calcula acessorios
    If iLarg > 0.8 Then qtPes = 6 Else qtPes = 4

    iAcess = vMaterial("pe") * qtPes * markup(2)

    vCxInf = Round(finalCxInf + lacaInf + iAcess, 2)

End Function

'valor das portas superiores
Function vPortaSup(largPS) As Single
    
    'variaveis
    Const ALTP As Single = 0.69
    Const LRIPA As Single = 0.05
    
    Dim baseP As Single
    Dim molduraP As Single
    Dim vFinal05 As Single
    Dim vFinal15 As Single
    Dim finalP As Single
    Dim m2Laca As Single
    Dim lacaP As Single
    Dim mEsp As Single
    Dim pDobs As Single
    Dim pPux As Single
    Dim pAcess As Single

    'calcula m2 da porta
    baseP = largPS * ALTP
    molduraP = (largPS + ALTP) * 2 * LRIPA
    
    vFinal05 = baseP * vMaterial(5)
    vFinal15 = molduraP * vMaterial(15)
    
    finalP = (vFinal05 + vFinal15) * markup(3)
    
          
    'calcula pintura da porta
    m2Laca = baseP + molduraP
    lacaP = m2Laca * vMaterial("laca")


    'calcula acessorios
    mEsp = ((largPS - LRIPA) * (ALTP - LRIPA)) * vMaterial("esp") + 10
    pDobs = 2 * vMaterial("dobs")
    pPux = 1 * vMaterial("pux")
    pAcess = (mEsp + pDobs + pPux) * markup(2)

    vPortaSup = Round(finalP + lacaP + pAcess, 2)

End Function

'valor das portas inferiores
Function vPortaInf(largPI, moldRpdAplqPI) As Single
    
    'variaveis
    Const ALTP As Single = 0.58
    Const LRIPA As Single = 0.05
    
    Dim baseP As Single
    Dim molduraP As Single
    Dim rpd As Single
    Dim m2Aplq As Single
    Dim aplq As Single
    Dim vFinal05 As Single
    Dim vFinal10 As Single
    Dim vFinal15 As Single
    Dim finalP As Single
    Dim m2Laca As Single
    Dim lacaP As Single
    Dim pDobs As Single
    Dim pPux As Single
    Dim pAcess As Single


    'calcula m2 da porta
    baseP = largPI * ALTP
    m2Aplq = (largPI - 0.17) * (ALTP - 0.17)

    If moldRpdAplqPI = "mold" Then
        molduraP = (largPI + ALTP) * 2 * LRIPA
        rpd = 0
        aplq = 0
    ElseIf moldRpdAplqPI = "rpd" Then
        rpd = Round((largPI / 0.03), 0) * ALTP * 0.015
        molduraP = 0
        aplq = 0
    ElseIf moldRpdAplqPI = "aplq" Then
        rpd = 0
        molduraP = (largPI + ALTP) * 2 * LRIPA
        aplq = m2Aplq
    End If
        
    'calcula valor das chapas usadas
    vFinal05 = (baseP + rpd) * vMaterial(5)
    vFinal10 = aplq * vMaterial(10)
    vFinal15 = molduraP * vMaterial(15)
    
    finalP = (vFinal05 + vFinal10 + vFinal15) * markup(3)
    
          
    'calcula pintura da porta
    m2Laca = baseP * 2
    lacaP = m2Laca * vMaterial("laca")


    'calcula acessorios
    pDobs = 2 * vMaterial("dobs")
    pPux = 1 * vMaterial("pux")
    pAcess = (pDobs + pPux) * markup(2)

    'retorna total
    vPortaInf = Round(finalP + lacaP + pAcess, 2)

End Function

'valor da gaveta
Function vGaveta(largG, moldRpdAplq, profG) As Single
    
    'variaveis
    Const ALTG As Single = 0.2
    Const LRIPA As Single = 0.05
    
    Dim caixaG As Single
    Dim cLarg As Single
    Dim cAlt As Single
    Dim cProf As Single
    Dim cFundo As Single
    Dim afastador As Single
    Dim baseG As Single
    Dim molduraG As Single
    Dim rpd As Single
    Dim aplq As Single
    Dim vFinal05 As Single
    Dim vFinal10 As Single
    Dim vFinal15 As Single
    Dim finalG As Single
    Dim m2Laca As Single
    Dim lacaG As Single
    Dim gCorr As Single
    Dim gPux As Single
    Dim gAcess As Single


    'calcula caixa da gaveta
    cLarg = largG - 0.02
    cProf = profG - 0.1
    cAlt = ALTG - 0.06
    cFundo = cLarg * cProf
    caixaG = ((cLarg + cProf) * 2 * cAlt)

    afastador = (profG - 0.05) * ALTG * 2


    'calcula m2 da frente
    baseG = largG * ALTG

    If moldRpdAplq = "mold" Then
        molduraG = (largG + ALTG) * 2 * LRIPA
        rpd = 0
        aplq = 0
    ElseIf moldRpdAplq = "rpd" Then
        rpd = Round((largG / 0.03), 0) * ALTG * 0.015
        molduraG = 0
        aplq = 0
    ElseIf moldRpdAplq = "aplq" Then
        rpd = 0
        molduraG = (largG + ALTG) * 2 * LRIPA
        aplq = 0
    End If
        

    'calcula valor das chapas usadas
    vFinal05 = (baseG + rpd + cFundo) * vMaterial(5)
    vFinal10 = aplq * vMaterial(10)
    vFinal15 = (molduraG + caixaG + afastador) * vMaterial(15)
    
    finalG = (vFinal05 + vFinal10 + vFinal15) * markup(3)
    
          
    'calcula pintura da gaveta
    m2Laca = (baseG * 2) + afastador
    lacaG = m2Laca * vMaterial("laca")


    'calcula acessorios
    gCorr = 1 * vMaterial("corr")
    gPux = 1 * vMaterial("pux")
    gAcess = (gCorr + gPux) * markup(2)


    'retorna total
    vGaveta = Round(finalG + lacaG + gAcess, 2)

End Function
    
'retorna valor total dos banheiros
Function vBanheiros(modelo, largS, largI, altS, altI, profS, profI, mold, qtPInf) As Single

    Const largP1Sup As Single = 0.48
    Const largP1Inf As Single = 0.33
    Const largP2Sup As Single = 0.4
    Const largP2Inf As Single = 0.28

    Dim qtGaveta As Integer
    Dim qtLatSup As Integer
    Dim qtRipaSup As Integer
    Dim qtLatInf As Integer
    Dim qtRipaInf As Integer
    

    Select Case modelo
        Case "Branco":
            qtLatSup = 3: qtRipaSup = 2
            qtLatInf = 2: qtRipaInf = 2: qtGaveta = 0
        Case "Verde":
            qtLatSup = 3: qtRipaSup = 0
            qtLatInf = 2: qtRipaInf = 0: qtGaveta = 0
        Case "Azul":
            qtLatSup = 4: qtRipaSup = 2
            qtLatInf = 3: qtRipaInf = 2: qtGaveta = 1
            If largI < 0.6 Then qtPInf = 1
        Case "Cinza":
            qtLatSup = 3: qtRipaSup = 0
            qtLatInf = 2: qtRipaInf = 0: qtGaveta = 0
    End Select

    vBanheiros = vCxSup(largS, altS, profS, qtLatSup, qtRipaSup) + vPortaSup(largP1Sup) + _
        vCxInf(largI, altI, profI, qtLatInf, qtRipaInf) + (vPortaInf(largP1Inf, mold) * qtPInf) + _
        (vGaveta(largP1Inf, mold, profI) * qtGaveta)

End Function

Function vNicho(largN, altN, profN) As Single

    Dim m2Nicho As Single
    Dim latsN As Single
    Dim basesN As Single

    ' m2
    basesN = (largN * 2) * 2
    latsN = (altN * 2) * 2
    m2Nicho = ((basesN + latsN) * profN) + (largN * altN)

    ' $
    vNicho = m2Nicho * vMaterial(15) * markup(3)
    
End Function

' TODO: Incluir nichos
'       Orçar opções de cuba
'       Tampo de granito



