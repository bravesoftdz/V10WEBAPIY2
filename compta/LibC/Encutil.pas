unit EncUtil;

interface

uses
  Windows, Messages, SysUtils, Classes,
  Graphics, ComCtrls, Controls, StdCtrls,
  Ent1, HEnt1, Hqry, hmsgbox, HCtrls,
{$IFDEF EAGLCLIENT}
  uTOB,
{$ELSE}
  DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
  SaisUtil,
  SaisComm,
  LettUtil ;

function  BonCompteEnc ( ENC : boolean ; Cpte : String ) : boolean ;
procedure DetruitGuideEnc ( Totale : boolean ) ;
procedure LGCommun ( Q : TQuery ; NumL : integer ) ;
procedure CATTOMP ( CAT : String3 ; it,va : HTStrings ; SorteLettre : TSorteLettre ; AvecTous : boolean ) ;
function  WhereGeneCritEcr ( Pages : TPageControl ; QEcr : THQuery ; QueEcr : boolean ) : String ;
function  WhereGeneCritDispo ( Pages : TPageControl ; QEcr : THQuery ; Pref : String) : String ;

implementation

function WhereGeneCritEcr ( Pages : TPageControl ; QEcr : THQuery ; QueEcr : boolean ) : String ;
Var StEcr : HString;
    Nam : String ;
    P   : TTabSheet ;
    C,C1: TControl ;
    i,j,k : integer ;
    StConf : String ;
BEGIN
for j:=0 to Pages.PageCount-1 do
    BEGIN
    P:=Pages.Pages[j] ;
    for i:=0 to P.ControlCount-1 do
        BEGIN
        C:=P.Controls[i] ;
        Nam:=C.Name ; if Nam[Length(Nam)]='_' then System.Delete(Nam,Length(Nam),1) ;
        if ((QueEcr) and (Copy(Nam,1,2)<>'E_')) then Continue ;
        if C is TGroupBox then
           begin
           for k:=0 to TGroupBox(C).ControlCount-1 do
              begin
              C1:=TGroupBox(C).Controls[k] ;
              Nam:=C1.Name ; if Nam[Length(Nam)]='_' then System.Delete(Nam,Length(Nam),1) ;
              if ((QueEcr) and (Copy(Nam,1,2)<>'E_')) then Continue ;
              QEcr.Control2Criteres(Nam,StEcr,C1,P) ;
              end ;
           end else QEcr.Control2Criteres(Nam,StEcr,C,P) ;
        END ;
    END ;
  if Copy(StEcr,2,3)='AND' then System.Delete(StEcr,1,5) ;
  if Copy(StEcr,2,2)='OR' then System.Delete(StEcr,1,4) ;
{$IFDEF EAGLCLIENT}
{$ELSE}
  StConf:=SQLConf('ECRITURE') ;
{$ENDIF}

  if StConf<>'' then
    StEcr:=StEcr+' AND '+StConf ;
  WhereGeneCritEcr:=StEcr ;
END ;

function WhereGeneCritDispo ( Pages : TPageControl ; QEcr : THQuery ; Pref : String) : String ;
Var StEcr : HString;
    Nam : String ;
    P   : TTabSheet ;
    C   : TControl ;
    i,j : integer ;
BEGIN
for j:=0 to Pages.PageCount-1 do
    BEGIN
    P:=Pages.Pages[j] ;
    for i:=0 to P.ControlCount-1 do
        BEGIN
        C:=P.Controls[i] ;
        Nam:=C.Name ; if Nam[Length(Nam)]='_' then System.Delete(Nam,Length(Nam),1) ;
        if (Pref='') then QEcr.Control2Criteres(Nam,StEcr,C,P)
                     else
                     BEGIN
                     if Copy(Nam,1,Pos('_',Nam))=Pref then QEcr.Control2Criteres(Nam,StEcr,C,P) ;
                     END ;
        END ;
    END ;
if Copy(StEcr,2,3)='AND' then System.Delete(StEcr,1,5) ; if Copy(StEcr,2,2)='OR' then System.Delete(StEcr,1,4) ;
WhereGeneCritDispo:=StEcr ;
END ;

procedure CATTOMP ( CAT : String3 ; it,va : HTStrings ; SorteLettre : TSorteLettre ; AvecTous : boolean ) ;
Var Q : TQuery ;
BEGIN
It.Clear ; Va.Clear ;
if CAT='' then
   BEGIN
   if SorteLettre=tslCheque then Q:=OpenSQL('Select MP_MODEPAIE, MP_LIBELLE from MODEPAIE Where MP_LETTRECHEQUE="X"',True) else
    if SorteLettre in [tslBOR,tslTraite]
         then Q:=OpenSQL('Select MP_MODEPAIE, MP_LIBELLE from MODEPAIE Where MP_LETTRETRAITE="X"',True)
         else Q:=OpenSQL('Select MP_MODEPAIE, MP_LIBELLE from MODEPAIE',True) ;
   END else
   BEGIN
   Q:=OpenSQL('Select MP_MODEPAIE, MP_LIBELLE from MODEPAIE Where MP_CATEGORIE="'+CAT+'"',True) ;
   END ;
if ((SorteLettre<>tslAucun) or (AvecTous)) then BEGIN It.Add(TraduireMemoire('<<Tous>>')) ; Va.Add('') ; END ;
While not Q.EOF do
  BEGIN
  Va.Add(Q.Fields[0].AsString) ; It.Add(Q.Fields[1].AsString) ;
  Q.Next ;
  END ;
Ferme(Q) ;
END ;

function BonCompteEnc ( ENC : boolean ; Cpte : String ) : boolean ;
Var Q : TQuery ;
    StE : String ;
BEGIN
Q:=OpenSQL('Select G_SUIVITRESO from GENERAUX Where G_GENERAL="'+Cpte+'"',True) ;
if Not Q.EOF then
   BEGIN
   StE:=Q.Fields[0].AsString ;
   if ENC then Result:=((StE='ENC') or (StE='MIX')) else Result:=((StE='DEC') or (StE='MIX')) ;
   END else Result:=False ;
Ferme(Q) ;
END ;

procedure DetruitGuideEnc ( Totale : boolean ) ;
BEGIN
if Totale then
   BEGIN
   ExecuteSQL('DELETE from GUIDE where GU_TYPE="'+V_PGI.USer+'"') ;
   ExecuteSQL('DELETE from ECRGUI where EG_TYPE="'+V_PGI.USer+'"') ;
   END else
   BEGIN
   ExecuteSQL('DELETE from GUIDE where GU_TYPE="'+V_PGI.USer+'" AND GU_GUIDE="'+V_PGI.User+'"') ;
   ExecuteSQL('DELETE from ECRGUI where EG_TYPE="'+V_PGI.USer+'" AND EG_GUIDE="'+V_PGI.User+'"') ;
   END ;
END ;

procedure LGCommun ( Q : TQuery ; NumL : integer ) ;
Var St : String ;
    i  : integer ;
BEGIN
Q.FindField('EG_TYPE').AsString:=V_PGI.User ;
Q.FindField('EG_GUIDE').AsString:=V_PGI.User ;
Q.FindField('EG_NUMLIGNE').AsInteger:=NumL ;
St:=Q.FindField('EG_ARRET').AsString ; for i:=1 to Length(St) do St[i]:='-' ; While Length(St)<10 do St:=St+'-' ;
Q.FindField('EG_ARRET').AsString:=St ;
Q.FindField('EG_DEBITDEV').AsString:='' ; Q.FindField('EG_CREDITDEV').AsString:='' ;
Q.FindField('EG_DEBITDEV').AsString:='' ; Q.FindField('EG_CREDITDEV').AsString:='' ;
END ;

end.
