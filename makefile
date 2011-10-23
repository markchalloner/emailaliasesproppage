#----- Include the PSDK's WIN32.MAK to pick up defines-------------------
!include <win32.mak>

PROJ = userproppage
all:    $(OUTDIR) $(OUTDIR)\$(PROJ).dll

$(OUTDIR):
     if not exist "$(OUTDIR)/$(NULL)" mkdir "$(OUTDIR)"

EXTRA_LIBS = dsprop.lib adsiid.lib activeds.lib odbccp32.lib odbc32.lib shell32.lib advapi32.lib

LINK32_OBJS = $(OUTDIR)\util.obj \
              $(OUTDIR)\userproppage.obj \
              $(OUTDIR)\stdafx.obj \
              $(OUTDIR)\proppageuserui.obj \
              $(OUTDIR)\proppageuser.obj \
              $(OUTDIR)\userproppage.res

CLOPT= /DWIN32 /D_WINDOWS /D_MBCS /D_WINDLL /D_USRDLL -D_AFX_NO_BSTR_SUPPORT /D_AFXDLL /DUNICODE /D_UNICODE

.cpp{$(OUTDIR)}.obj:
     $(cc) $(cflags) $(cdebug) $(cvarsdll) /GR $(CLOPT) /EHsc /D_AFXDLL /D_USRDLL /D_MBCS /Fo"$(OUTDIR)\\" /Fd"$(OUTDIR)\\" $**

$(OUTDIR)\userproppage.res:
     $(rc) $(rcflags) $(rcvars) /Fo$(OUTDIR)\userproppage.res userproppage.rc

$(OUTDIR)\$(PROJ).dll:   $(LINK32_OBJS)
     $(link) $(ldebug) $(dlllflags) /MACHINE:$(CPU) /PDB:$(OUTDIR)\$(PROJ).pdb -out:$(OUTDIR)\$(PROJ).dll /DEF:userproppage.def  $(LINK32_OBJS) $(EXTRA_LIBS) $(baselibs) $(olelibs) 

clean:
     $(CLEANUP)
