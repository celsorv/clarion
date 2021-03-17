!***************
!* Nome      : DeleteFiles v1.00
!* Descrição : Apaga arquivos de um diretório
!* Autor     : Celso R. Vitorino
!* Data      : 26/02/2002
!* Parâmetros: Diretório, nome de arquivo (aceita coringa)
!* Retorno   : True, False
!***************
!
DeleteFiles FUNCTION(SrcFolder,SrcFiles)

i                   LONG,AUTO
loc:Sucess          BYTE,AUTO
loc:CountFiles      LONG,AUTO
QFiles              QUEUE(File:queue),PRE(FIL).

  CODE

    loc:Sucess = True

    Directory(QFiles,Clip(SrcFolder) & SrcFiles,ff_:NORMAL)
    loc:CountFiles = Records(QFiles)

    Loop i = 1 To loc:CountFiles
      Get(QFiles,i)
      Remove(Clip(SrcFolder) & FIL:ShortName)
      if ErrorCode()
        loc:Sucess = False
        Break
      End
    End

    Return(loc:Sucess)
    