!***************
!*
!*  Data     : 26/02/2002
!*  Autor    : Celso R. Vitorino
!*  Objetivo : Mostra relatório texto na tela
!*  Linguagem: Softvelocity - Clarion 4b For Windows
!*
!***************
!
MostraRelatorio   PROCEDURE

Loc:Contagem      LONG
Loc:NumPagina     LONG
Loc:TotalPagina   LONG
Loc:ReturnValue   BYTE

Que:Paginas       QUEUE
Posicao             STRING(4)
                  END

Que:Vitrine       QUEUE
Linha               STRING(127)
                  END

Preview:Window WINDOW('Visualização de Relatório'),AT(,,,),FONT('Arial',10,,),COLOR(0E3FFFFH), |
                 IMM,ICON(Icon:Print),STATUS(-1,160),GRAY,MAX,RESIZE,MAXIMIZE,ALRT(PgUpKey),ALRT(PgDnKey)
                 MENUBAR
                   MENU('&Arquivo'),USE(?Arquivo)
                     ITEM('&Imprimir'),USE(?ArquivoImprimir),MSG('Imprime o relatório')
                     ITEM,SEPARATOR
                     ITEM('Sai&r'),USE(?ArquivoSair),MSG('Sai sem imprimir o relatório'),STD(STD:Close)
                   END
                   MENU('&Visualização'),USE(?Visualizacao)
                     ITEM('&Próxima Página'),USE(?Next),MSG('Visualiza a próxima página do relatório')
                     ITEM('Página &Anterior'),USE(?Previous),DISABLE,MSG('Visualiza a página anterior do relatório')
                   END
                 END
                 TOOLBAR,AT(0,0,341,18)
                   BUTTON,AT(3,2,14,14),TIP('Imprime este relatório'),USE(?PrintButton),ICON(ICON:Print)
                   BUTTON,AT(19,2,14,14),TIP('Sai sem imprimir o relatório'),USE(?ExitButton),ICON(ICON:NoPrint)
                   PROMPT('&Página:'),AT(47,4),USE(?Loc:NumPagina:Prompt)
                   SPIN(@n8),AT(73,3,33,12),USE(Loc:NumPagina),RANGE(1,10), |
                     STEP(1),MSG('Selecione a página a visualizar')
                 END
                 LIST,AT(0,0),USE(?List1),HVSCROLL,FROM(Que:Vitrine), |
                   FONT('Courier',10,,FONT:regular),FULL
               END

  CODE
  Loc:ReturnValue=False
  OPEN(Preview:Window)
  TARGET{Prop:MinWidth}=315
  SELECT(?List1)
  !
  ! Carrega lista de páginas
  !
  Loc:Contagem=0
  SET(IMPRESSAO)
  LOOP
    NEXT(IMPRESSAO)
    IF ERRORCODE()
      BREAK
    END
    IF Loc:Contagem=0
      Que:Paginas.Posicao=POSITION(IMPRESSAO)
      ADD(Que:Paginas)
    END
    Loc:Contagem=Loc:Contagem+1
    IF Loc:Contagem=Glo:TamPagina
      Loc:Contagem=0
    END
  END
  Loc:NumPagina=1
  Loc:TotalPagina=RECORDS(Que:Paginas)
  IF Loc:TotalPagina<=1
    DISABLE(?Loc:NumPagina:Prompt)
    DISABLE(?Loc:NumPagina)
    DISABLE(?Visualizacao)
  ELSE
    ?Loc:NumPagina{Prop:RangeHigh}=Loc:TotalPagina
    ?Loc:NumPagina{Prop:Msg}='Informe um número de página de 1 a ' & Loc:NumPagina
  END
  DO CarregaPagina
  ACCEPT
    CASE EVENT()
    OF Event:NewSelection
      CASE FIELD()
      OF ?Loc:NumPagina
        DO CarregaPagina
      END
    OF Event:Rejected
      CASE REJECTCODE()
      OF Reject:RangeHigh
        CHANGE(FIELD(),FIELD(){Prop:RangeHigh})
      OF Reject:RangeLow
        CHANGE(FIELD(),FIELD(){Prop:RangeLow})
      END
      DO CarregaPagina
    OF Event:AlertKey
      CASE KEYCODE()
      OF PgUpKey
        POST(Event:Accepted,?Previous)
      OF PgDnKey
        POST(Event:Accepted,?Next)
      END
    OF Event:Accepted
      CASE FIELD()
      OF ?ArquivoSair OROF ?ExitButton
        POST(Event:CloseWindow)
      OF ?Next
        IF Loc:NumPagina<Loc:TotalPagina
          Loc:NumPagina=(Loc:NumPagina+1)
          DO CarregaPagina
        END
      OF ?Previous
        IF Loc:NumPagina>1
          Loc:NumPagina=(Loc:NumPagina-1)
          DO CarregaPagina
        END
      OF ?ArquivoImprimir OROF ?PrintButton
        POST(Event:CloseWindow)
        Loc:ReturnValue=True
      OF ?Loc:NumPagina
        DO CarregaPagina
      END
    OF Event:Maximized OROF Event:Sized
      ?List1{Prop:Format}=''
    END
  END
  CLOSE(Preview:Window)
  IF Loc:ReturnValue
    GlobalResponse=RequestCompleted
  ELSE
    GlobalResponse=RequestCancelled
  END

  RETURN

CarregaPagina ROUTINE

  FREE(Que:Vitrine)

  GET(Que:Paginas,Loc:NumPagina)
  IF ERRORCODE() THEN EXIT.
  RESET(IMPRESSAO,Que:Paginas.Posicao)
  Loc:Contagem=0
  LOOP
    NEXT(IMPRESSAO)
    IF ERRORCODE()
      BREAK
    END
    Que:Vitrine.Linha=IMPRESSAO.LINHA
    ADD(Que:Vitrine)
    Loc:Contagem=Loc:Contagem+1
    IF Loc:Contagem=Glo:TamPagina
      BREAK
    END
  END

  IF Loc:TotalPagina>1
    IF Loc:NumPagina=1
      DISABLE(?Previous)
    ELSE
      ENABLE(?Previous)
    END
    IF Loc:NumPagina=Loc:TotalPagina
      DISABLE(?Next)
    ELSE
      ENABLE(?Next)
    END
  END
  TARGET{Prop:StatusText,2}='Pagina ' & Loc:NumPagina & ' de ' & Loc:TotalPagina
  DISPLAY

!-------------------------------------------------------
! Retorna uma String sem acentos
!-------------------------------------------------------
SemAcentos FUNCTION(STRING Texto)

I                LONG
J                BYTE
Tam              SHORT
TextoRetorno     STRING(300)
LetrasAcentuadas STRING(4),Dim(14,2)

  CODE
    LetrasAcentuadas[01,1] = 'ÁÀÃÂ'
    LetrasAcentuadas[01,2] = 'A'
    LetrasAcentuadas[02,1] = 'Ç'
    LetrasAcentuadas[02,2] = 'C'
    LetrasAcentuadas[03,1] = 'ÉÈÊ'
    LetrasAcentuadas[03,2] = 'E'
    LetrasAcentuadas[04,1] = 'Í'
    LetrasAcentuadas[04,2] = 'I'
    LetrasAcentuadas[05,1] = 'Ñ'
    LetrasAcentuadas[05,2] = 'N'
    LetrasAcentuadas[06,1] = 'ÓÕÔÖ'
    LetrasAcentuadas[06,2] = 'O'
    LetrasAcentuadas[07,1] = 'ÚÜ'
    LetrasAcentuadas[07,2] = 'U'
  
    LetrasAcentuadas[08,1] = 'áàãâ'
    LetrasAcentuadas[08,2] = 'a'
    LetrasAcentuadas[09,1] = 'ç'
    LetrasAcentuadas[09,2] = 'c'
    LetrasAcentuadas[10,1] = 'éèê'
    LetrasAcentuadas[10,2] = 'e'
    LetrasAcentuadas[11,1] = 'í'
    LetrasAcentuadas[11,2] = 'i'
    LetrasAcentuadas[12,1] = 'ñ'
    LetrasAcentuadas[12,2] = 'n'
    LetrasAcentuadas[13,1] = 'óõôö'
    LetrasAcentuadas[13,2] = 'o'
    LetrasAcentuadas[14,1] = 'úü'
    LetrasAcentuadas[14,2] = 'u'

    Tam=SIZE(Texto)
       
    TextoRetorno=Texto
    LOOP I=1 TO Tam
      LOOP J=1 TO 14
        IF INSTRING(TextoRetorno[I],CLIP(LetrasAcentuadas[J,1]))
          TextoRetorno[I]=LetrasAcentuadas[J,2]
          BREAK
        END
      END
    END

    RETURN(TextoRetorno[1 : Tam])

!-------------------------------------------------------
! Grava registro do relatorio
!-------------------------------------------------------
GravaLinha  PROCEDURE(STRING Texto)

  CODE
    IMPRESSAO.LINHA=SemAcentos(Texto)
    ADD(IMPRESSAO)
    
    RETURN

!------------------------------------------------------
! Impressao da linha-detalhe do relatorio
!------------------------------------------------------
Imprime     PROCEDURE(<STRING Texto>, <SHORT EmBranco>)

I  BYTE

  CODE
    IF OMITTED(1) THEN Texto=' '.
    IF OMITTED(2) THEN EmBranco=0.

!------ Quebra a pagina
    
    IF GLO:QuebraPagina
      IF GLO:TamCabecalho
        AbrePagina()
        EmBranco=CHOOSE(EmBranco<0,0,EmBranco)
      END
      GLO:QuebraPagina=False
    END

!------ Imprime linhas em branco antes da linha detalhe -----------
    
    IF EmBranco<0
      I=ABS(EmBranco)
      LOOP I TIMES
        IF GLO:TamCabecalho
          IF GLO:CabContador>(GLO:TamPagina-7)
            AbrePagina()
            BREAK
          END
        END
        GravaLinha(' ')
        GLO:CabContador=(GLO:CabContador+1)
      END
    END

!------ Imprime linha detalhe -------------

    IF GLO:TamCabecalho
      IF GLO:CabContador>(GLO:TamPagina-7)
        AbrePagina()
      END
    END
    
    GravaLinha(Texto)
    GLO:CabContador=(GLO:CabContador+1)

!------ Imprime linhas em branco depois da linha detalhe ---------
    
    IF EmBranco>0
      I=EmBranco
      LOOP I TIMES
        IF GLO:TamCabecalho
          IF GLO:CabContador>(GLO:TamPagina-7)
            AbrePagina()
            BREAK
          END
        END
        GravaLinha(' ')
        GLO:CabContador=(GLO:CabContador+1)
      END
    END

    RETURN
    
!----------------------------------------------------
! Impressao do cabecalho do relatorio
!----------------------------------------------------
AbrePagina     PROCEDURE(<SHORT Modo>)

I  BYTE

  CODE

    IF OMITTED(1) THEN Modo=0.

!------ Imprime linhas de rodape --------------

    IF GLO:TamCabecalho    
      IF GLO:CabPagina
        I=((GLO:TamPagina-6)-GLO:CabContador)
        LOOP I TIMES
          GravaLinha(' ')
        END
        GravaLinha(' ')
        GravaLinha('Emitido em: '&FORMAT(TODAY(),@D06B)&ALL(' ',73)& |
                   GLO:NomeVers&' - Pagina: '&FORMAT(GLO:CabPagina,@N_5))
        LOOP 3 TIMES
          GravaLinha(' ')
        END
        IF Modo
          GravaLinha(Prn:10cpp)
        ELSE
          GravaLinha(' ')
        END
        GLO:CabContador=GLO:TamPagina
      END
    END
        
    IF Modo THEN RETURN.
    
!------ Imprime linhas do cabecalho -----------
    
    IF GLO:TamCabecalho
      IF GLO:CabEstilo
        GravaLinha(Prn:12cpp)   !Comprimido
      ELSE
        GravaLinha(Prn:10cpp)   !Normal
      END
      GravaLinha(' ')
      GravaLinha(' ')
      LOOP I=1 TO GLO:TamCabecalho
        GravaLinha(GLO:LinCabecalho[I])
      END
      GLO:CabContador=(GLO:TamCabecalho+3)
    END

    GLO:CabPagina=(GLO:CabPagina+1)

    RETURN

!----------------------------------------------------
! Retorna uma string de espacos
!----------------------------------------------------
x     FUNCTION(SHORT Tam)

  CODE
    RETURN(ALL(' ',Tam))


!----------------------------------------------------
! Retorna uma string duplicada
!----------------------------------------------------
r     FUNCTION(STRING Texto,SHORT Tam)

I        SHORT
J        SHORT
Taman    SHORT
TXTConv  STRING(1024)

  CODE
    Taman=LEN(Texto)
    
    I=(Tam-1)
    J=Taman
    TXTConv=Texto
  
    LOOP I TIMES
      TXTConv=TXTConv[1 : J] & Texto
      J=(J+Taman)
    END

   RETURN(TXTConv[1 : J])

!-------------------------------------------
! Configura exportação de relatório para DBF
!-------------------------------------------
ConfigRelDBF     PROCEDURE

Loc:RetArq        STRING(50)
Loc:NmDBF         STRING(50)

config:window WINDOW('Configurar Arquivo DBF'),AT(,,221,55),FONT('Arial',9,,),CENTER,GRAY,DOUBLE
                GROUP('Arquivo a Gravar'),AT(7,7,146,27),USE(?Arquivo:Group),BOXED
                  ENTRY(@s50),AT(13,17,,10),USE(Loc:NmDBF),SKIP,READONLY
                  BUTTON,AT(134,16,13,12),USE(?ArqDBFButton),ICON(ICON:Open)
                END
                BUTTON('OK'),AT(164,10,48,14),USE(?OKButton),DEFAULT
                BUTTON('Cancelar'),AT(164,30,48,15),USE(?CancelButton)
              END

  CODE
  OPEN(Config:Window)
  SELECT(?Arquivo:Group)
  ACCEPT
    CASE EVENT()
    OF Event:Accepted
      CASE FIELD()
      OF ?OKButton
        IF ~Loc:NmDBF
          BEEP
          MESSAGE('Informe o local e nome do arquivo a gravar','Arquivo DBF',ICON:EXCLAMATION)
          POST(Event:Accepted,?ArqDBFButton)
          CYCLE
        END
        Loc:RetArq=Loc:NmDBF
        POST(Event:CloseWindow)
      OF ?CancelButton
        POST(Event:CloseWindow)
      OF ?ArqDBFButton
        IF FILEDIALOG('Selecione o Arquivo a Gravar',Loc:NmDBF,'Arquivos DBF|*.dbf|Todos os Arquivos|*.*',FILE:Save+FILE:KeepDir)
          IF NOT INSTRING('.',Loc:NmDBF)
            Loc:NmDBF=CLIP(Loc:NmDBF)&'.dbf'
          END        
          DISPLAY(?Loc:NmDBF)
          SELECT(?OkButton)
        END
      END
    END
  END
  CLOSE(Config:Window)
  IF Loc:RetArq
    GlobalResponse=RequestCompleted
  ELSE
    GlobalResponse=RequestCancelled
  END

  RETURN(Loc:RetArq)

!---------------------------------------
! Configura impressão de relatório texto
!---------------------------------------
ConfigRelatorio   PROCEDURE

Loc:ReturnValue   BYTE

Config:Window WINDOW('Configurar impressão texto'),AT(,,236,96),CENTER,IMM,GRAY,DOUBLE
                OPTION('&Estilo de Impressão'),AT(6,4,158,33),USE(Glo:CabEstilo),BOXED
                  RADIO('Normal'),AT(13,19,38,12),USE(?Normal:Radio),VALUE('0')
                  RADIO('Comprimido'),AT(69,19,51,12),USE(?Comprimido:Radio),VALUE('1')
                END
                GROUP('&Imagem da Impressão'),AT(6,45,223,45),USE(?Imagem:Group),BOXED
                  CHECK('&Criar imagem da impressão?'),AT(13,57),USE(Glo:CabGrvImagem)
                  PROMPT('Arquivo:'),AT(42,71),USE(?Glo:CabArqImagem:Prompt),DISABLE
                  ENTRY(@s50),AT(72,71,134,10),USE(Glo:CabArqImagem),DISABLE,READONLY
                  BUTTON,AT(209,70,13,12),USE(?ArqImagem:Button),DISABLE,SKIP,ICON(ICON:Open)
                END
                BUTTON('OK'),AT(181,8,48,14),USE(?OkButton),DEFAULT
                BUTTON('Cancelar'),AT(181,26,48,14),USE(?CancelButton)
              END

  CODE
  CLEAR(Glo:CabGrvImagem)
  CLEAR(Glo:CabArqImagem)
  Loc:ReturnValue=False
  OPEN(Config:Window)
  SELECT(?Imagem:Group)
  ACCEPT
    CASE EVENT()
    OF Event:Accepted
      CASE FIELD()
      OF ?OkButton
        IF Glo:CabGrvImagem AND ~Glo:CabArqImagem
          BEEP
          MESSAGE('Informe o local e nome do arquivo a gravar','Arquivo de imagem',ICON:EXCLAMATION)
          POST(Event:Accepted,?ArqImagem:Button)
          CYCLE
        END
        Loc:ReturnValue=True
        POST(Event:CloseWindow)
      OF ?CancelButton
        POST(Event:CloseWindow)
      OF ?Glo:CabGrvImagem
        CASE Glo:CabGrvImagem
        OF 1
          ENABLE(?Glo:CabArqImagem:Prompt)
          ENABLE(?Glo:CabArqImagem)
          ENABLE(?ArqImagem:Button)
          POST(Event:Accepted,?ArqImagem:Button)
        OF 0
          CLEAR(Glo:CabArqImagem)
          DISPLAY(?Glo:CabArqImagem)
          DISABLE(?Glo:CabArqImagem:Prompt)
          DISABLE(?Glo:CabArqImagem)
          DISABLE(?ArqImagem:Button)
        END
      OF ?ArqImagem:Button
        IF FILEDIALOG('Selecione o Arquivo a Gravar',Glo:CabArqImagem,'Arquivos Texto|*.txt|Todos os Arquivos|*.*',FILE:Save+FILE:KeepDir)
          IF NOT INSTRING('.',Glo:CabArqImagem)
            Glo:CabArqImagem=CLIP(Glo:CabArqImagem)&'.txt'
          END        
          DISPLAY(?Glo:CabArqImagem)
          SELECT(?Glo:CabArqImagem)
        END
      END
    END
  END
  CLOSE(Config:Window)
  IF Loc:ReturnValue
    GlobalResponse=RequestCompleted
  ELSE
    GlobalResponse=RequestCancelled
  END

  RETURN

!ENVIA RELATORIO PARA IMPRESSORA
!------------------------------------------------------------------------------
SendRelatorio      PROCEDURE(STRING Texto)

  CODE
    RUN('prnsp.exe f="' & CLIP(Glo:CabArquivo) & '" p="' & |
        Printer{PropPrint:Device} & '" d="' & CLIP(Texto) & '"')

    IF RUNCODE()
      IF RUNCODE()=-4
        MESSAGE(ERROR(),'Erro de Printer Spool',ICON:EXCLAMATION)
      ELSE
        MESSAGE('Código de retorno: ' & RUNCODE(),'Erro de Printer Spool',ICON:EXCLAMATION)
      END
    END

    RETURN
    