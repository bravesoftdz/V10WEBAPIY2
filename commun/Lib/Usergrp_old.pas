unit Usergrp;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  FichList, DBCtrls, StdCtrls, Mask, hmsgbox, DB, DBTables, Buttons, UserAcc,
  ExtCtrls, Grids, DBGrids, HDB, Ent1, HEnt1, HCtrls, Spin, HSysMenu,
  HTB97,MajTable, Hqry, HPanel, UiUtil ;

Procedure FicheUserGrp ;

type
  TFusergrp = class(TFFicheListe)
    PJournaux: TPanel;
    EJOURNALAUTORISE: TLabel;
    PFenGuide: TPanel;
    Panel2: TPanel;
    BCValide: TToolbarButton97;
    BCAbandon: TToolbarButton97;
    Bdetag: TToolbarButton97;
    BTag: TToolbarButton97;
    JOURNALAUTORISE: TListBox;
    UG_GROUPE: TDBEdit;
    TUG_GROUPE: THLabel;
    TUG_LIBELLE: THLabel;
    TUG_PASSWORD: THLabel;
    TUG_JALAUTORISES: TLabel;
    UG_JALAUTORISES: TDBEdit;
    UG_PASSWORD: TDBEdit;
    UG_LIBELLE: TDBEdit;
    UG_ABREGE: TDBEdit;
    TUG_ABREGE: THLabel;
    UG_NUMERO: TDBEdit;
    TUG_NUMERO: THLabel;
    ZoomJAL: TToolbarButton97;
    TaUG_GROUPE: TStringField;
    TaUG_LIBELLE: TStringField;
    TaUG_ABREGE: TStringField;
    TaUG_CONFIDENTIEL: TStringField;
    TaUG_PASSWORD: TStringField;
    TaUG_NUMERO: TIntegerField;
    TaUG_MONTANTMIN: TFloatField;
    TaUG_MONTANTMAX: TFloatField;
    TaUG_JALAUTORISES: TStringField;
    BAcces: TToolbarButton97;
    TaUG_NIVEAUACCES: TIntegerField;
    TaUG_LANGUE: TStringField;
    TaUG_PERSO: TStringField;
    GBMontant: TGroupBox;
    Label1: THLabel;
    UG_MONTANTMIN: TDBEdit;
    Label2: THLabel;
    UG_MONTANTMAX: TDBEdit;
    UG_LANGUE: THDBValComboBox;
    tLangue: THLabel;
    tPerso: THLabel;
    UG_PERSO: THDBValComboBox;
    TDC_NIVEAUACCES: THLabel;
    UG_NIVEAUACCES: THDBSpinEdit;
    UG_CONFIDENTIEL: TDBCheckBox;
    procedure TaBeforeDelete(DataSet: TDataSet);
    procedure TaAfterDelete(DataSet: TDataSet);
    procedure FormShow(Sender: TObject);
    procedure TaAfterPost(DataSet: TDataSet);
    procedure BAccesClick(Sender: TObject);
    procedure BCValideClick(Sender: TObject);
    procedure BCAbandonClick(Sender: TObject);
    procedure ZoomJALClick(Sender: TObject);
    procedure BTagClick(Sender: TObject);
    procedure BdetagClick(Sender: TObject);
    procedure UG_JALAUTORISESExit(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure UG_LANGUEClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure UG_MONTANTMINExit(Sender: TObject);
    procedure UG_MONTANTMINEnter(Sender: TObject);
    procedure UG_MONTANTMAXEnter(Sender: TObject);
    procedure UG_MONTANTMAXExit(Sender: TObject);
    procedure STaDataChange(Sender: TObject; Field: TField);
  private    { Déclarations privées }
    NextNumGrp  : Integer ;
    MemoLgu : String ;
    TJAL : TStringList ;
    Function  EnregOK : boolean ; Override ;
    Procedure NewEnreg ; Override ;
    Procedure ChargeEnreg ; Override ;
    Procedure TrouveTrou ;
    Procedure AutoriseNouveau ;
    Procedure RempliJAL ;
    Procedure TagDetag(Ok:Boolean) ;
    Procedure MajSelect ;
    function  VerifCodeJal : boolean ;
    Procedure ChargMagUserGrp ;
  public    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Procedure FicheUserGrp ;
var Fusergrp: TFusergrp;
    PP : THPanel ;
BEGIN
Fusergrp:=TFusergrp.Create(Application) ;
Fusergrp.InitFL('UG','PRT_USERGRP','','',taModif,TRUE,Fusergrp.TaUG_GROUPE,
                Fusergrp.TaUG_LIBELLE,Fusergrp.TaUG_GROUPE,['ttUserGroupe']) ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     Fusergrp.ShowModal ;
    Finally
     Fusergrp.Free ;
    End ;
   Screen.Cursor:=crDefault ;
   END else
   BEGIN
   InitInside(FUserGrp,PP) ;
   FUserGrp.Show ;
   END ;
END ;

Procedure TFusergrp.TrouveTrou  ;
Var Q : TQuery ;
BEGIN
Q:=OpenSQL('Select UG_NUMERO from USERGRP order by UG_NUMERO',TRUE) ; NextNumGrp:=1 ;
While(Not Q.EOF)AND(Q.Fields[0].AsInteger=NextNumGrp) do BEGIN  Q.Next ; Inc(NextNumGrp) ;  END ;
if(Q.EOF)AND(NextNumGrp>100) then NextNumGrp:=0 ;
Ferme(Q) ;
END ;

Procedure TFusergrp.ChargeEnreg ;
BEGIN
 Inherited ;
if V_PGI.Groupe=TaUG_GROUPE.AsString then UG_PASSWORD.PasswordChar:=#0
                                  else UG_PASSWORD.PasswordChar:='*' ;
if V_PGI.Superviseur<>True then FicheReadOnly(Self) ;
if MemoLgu<>UG_LANGUE.Value then UG_LANGUEClick(nil) ;
if TaUG_PERSO.AsString='' then UG_PERSO.ItemIndex:=UG_PERSO.Items.IndexOf(HM2.Mess[1])
                          else UG_PERSO.ItemIndex:=UG_PERSO.Values.Indexof(TaUG_PERSO.AsString) ;
MemoLgu:=TaUG_LANGUE.AsString ;
END ;

Procedure TFusergrp.NewEnreg ;
BEGIN
InHerited ;
BAcces.Enabled:=False ;
TaUG_NUMERO.AsInteger:=NextNumGrp ;
END ;

Function TFusergrp.EnregOK : boolean ;
BEGIN
result:=InHerited EnregOK ; if Not Result then Exit ;
Modifier:=True ; Result:=False ;
if Ta.state in [dsEdit,dsInsert] then
   BEGIN
//   if TaUG_ABREGE.AsString='' then BEGIN HM.Execute(6,'','') ; UG_ABREGE.SetFocus ; Exit ; END ;
//   if TaUG_PASSWORD.AsString='' then BEGIN HM.Execute(7,'','') ; UG_PASSWORD.SetFocus ; Exit ; END ;
   if not VerifCodeJal then Exit ;
   END ;
result:=TRUE  ; Modifier:=False ; ChargMagUserGrp ;
END ;

procedure TFusergrp.TaBeforeDelete(DataSet: TDataSet);
Var Q : TQuery ;
    i,j : Byte ;
    joker,St : String ;
    StLike : String[100] ;
    StLike1 : String[100] ;
begin
  inherited;
i:=TaUG_NUMERO.AsInteger ; StLike:='' ; StLike1:='' ;
for j:=1 to i-1 do StLike:=StLike+'_' ;
StLike1:=StLike ; StLike:=StLike+'X' ; StLike1:=StLike1+'-' ;
if V_PGI.Driver=dbMSACCESS then Joker:='*' else Joker:='%' ;
Q:=OpenSQL('Select MN_ACCESGRP from MENU Where MN_ACCESGRP Like "'+StLike+Joker+'" Or MN_ACCESGRP Like "'+StLike1+Joker+'"',False) ;
Q.UpdateMode:=upWhereChanged ;
While Not Q.EOF do
  BEGIN
  Q.Edit ; St:=Q.Fields[0].AsString ; St[i]:='0' ; Q.Fields[0].AsString:=St ; Q.Post ; Q.Next ;
  END ;
Ferme(Q) ;
ExecuteSql('Delete From UTILISAT Where US_GROUPE="'+UG_GROUPE.Text+'"') ;
end;

procedure TFusergrp.TaAfterDelete(DataSet: TDataSet);
begin
  inherited;
AutoriseNouveau ;
end;

procedure TFusergrp.FormShow(Sender: TObject);
begin
MemoLgu:='' ;
  inherited;
if ((EstSerie(S3)) or (EstSerie(S5))) then
   BEGIN
   tLangue.Visible:=False ; UG_Langue.Visible:=False ;
   tPerso.Visible:=False  ; UG_Perso.Visible:=False ;
   if EstSerie(S3) then
      BEGIN
      UG_CONFIDENTIEL.Visible:=False ;
      UG_NIVEAUACCES.Visible:=False ;
      TDC_NIVEAUACCES.Visible:=False ;
      END ;
   END ;
if ctxMode in V_PGI.PGIContexte then
   BEGIN
   tLangue.Visible:=True ; UG_Langue.Visible:=True ;
   tPerso.Visible:=True  ; UG_Perso.Visible:=True ;
   END ;
if (ctxGescom in V_PGI.PGIContexte)
   or (ctxAffaire in V_PGI.PGIContexte)
   // or ctxPaie in V_PGI.PGIContexte
   then  
   BEGIN
   UG_JALAUTORISES.Visible:=False ;
   TUG_JALAUTORISES.Visible:=False ;
   GBMontant.Visible:=False ;
   ZoomJAL.Visible:=False ;  
   END ;
RempliJAL ;
TaUG_MONTANTMIN.DisplayFormat:=StrfMask(V_PGI.OkDecV,'',True) ;
TaUG_MONTANTMAX.DisplayFormat:=TaUG_MONTANTMIN.DisplayFormat ;
AutoriseNouveau ;
if V_PGI.Superviseur<>True then FicheReadOnly(Self) else
  if(Ta.Eof) And (Ta.Bof) And (FTypeAction<>taConsult)then BinsertClick(Nil) ;
end;

procedure TFusergrp.TaAfterPost(DataSet: TDataSet);
begin
  inherited;
AutoriseNouveau ;
end;

Procedure TFusergrp.AutoriseNouveau ;
BEGIN
TrouveTrou ; BInsert.Enabled:=(NextNumGrp>0) ; BAcces.Enabled:=(V_PGI.Superviseur=True) ;
END ;

procedure TFusergrp.BAccesClick(Sender: TObject);
Var ii : integer ;
    SS : SetInteger ;
begin
  inherited;
if ctxAffaire in V_PGI.PGIContexte then
   BEGIN
   if ctxScot in V_PGI.PGIContexte then SS:=[141,142,143,144,151,152,154,26,27]
                                   else SS:=[71,72,73,138,139,92,74,140,26,27];
   END else
if ctxMode in V_PGI.PGIContexte then
   BEGIN
   if ctxFO in V_PGI.PGIContexte then SS:=[107,108,109,110,111,112,26]
   else SS:=[101,102,103,104,105,106,107,108,109,110,111,112,113,114,115,26];
   END else
if ctxGescom in V_PGI.PGIContexte then
   BEGIN
   SS:=[30,31,32,33,70,65,60,27,26];
   END else
if ctxPaie in V_PGI.PGIContexte then
   BEGIN
   SS:=[41,42,43,44,46,49] ;
   END else
   BEGIN
   if ctxPCL in V_PGI.PGIContexte then SS:=[52,53,54,55,56,96,26,27] else
     begin
     if EstSerie(S7) then SS:=[1,2,3,6,7,15,25,26] else
      if EstSerie(S5) then SS:=[6,9,11,12,13,14,16,17,18,20,26,27] else
       if EstSerie(S3) then SS:=[6,9,11,13,14,16,17,18,20,26,27] ;
     end;
   END ;

if ctxGRC in V_PGI.PGIContexte then
   BEGIN
   SS:=SS+[92];
   END;
   
// lanceur = tous les menus des applications multi-dossiers
// ok si AGL >= v504c
if ctxDP in V_PGI.PGIContexte then
   BEGIN                     
   SS:=[75,76,78]; // DP
   // compta
   if ctxPCL in V_PGI.PGIContexte then SS:=SS+[52,53,54,55,56,96,26,27] else
     begin
     if EstSerie(S7) then SS:=SS+[1,2,3,6,7,15,25,26] else
      if EstSerie(S5) then SS:=SS+[6,9,11,12,13,14,16,17,18,20,26,27] ;
     end;
   SS:=SS+[40,41,42,44,49]; // paie
   SS:=SS+[141,142,143,144,151,152,154]; // gestion interne sauf 26,27 déjà dans compta
   END ;
{$IFDEF CEGID}
ii:=FicheUserAcces(TaUG_GROUPE.AsString,TaUG_LIBELLE.AsString,TaUG_NUMERO.AsInteger,SS,'',False) ;
{$ELSE}
ii:=FicheUserAcces(TaUG_GROUPE.AsString,TaUG_LIBELLE.AsString,TaUG_NUMERO.AsInteger,SS) ;
{$ENDIF}
end;

procedure TFusergrp.BCValideClick(Sender: TObject);
var i : integer ;
    St : String ;
    Tous : Boolean ;
begin
  inherited;
BCAbandonClick(Nil) ;
if Ta.state in [dsInactive,dsBrowse] then Ta.Edit ;
St:='' ;
Tous:=True ;
for i:=0 to JOURNALAUTORISE.Items.Count-1 do
  BEGIN
  if JOURNALAUTORISE.Selected[i] then St:=St+TJAL.Strings[i]+';' else Tous:=False ;
  END ;
if not Tous then UG_JALAUTORISES.Text:=St else UG_JALAUTORISES.Text:='' ;
end;

procedure TFusergrp.BCAbandonClick(Sender: TObject);
begin
  inherited;
PAppli.Enabled:=True ; FListe.Enabled:=True ; PBouton.Enabled:=True ;
PJournaux.Visible:=False ;
UG_JALAUTORISES.SetFocus ;
end;

procedure TFusergrp.ZoomJALClick(Sender: TObject);
begin
  inherited;
if not VerifCodeJal then Exit ;
MajSelect ;
PJournaux.Visible:=True ; PJournaux.BringToFront ;
PAppli.Enabled:=False ; FListe.Enabled:=False ; PBouton.Enabled:=False ;
PJournaux.Left:=FListe.Width+(FListe.Width-PJournaux.Width) div 2 ;
PJournaux.Top:=PAppli.Top+(PAppli.Height-PJournaux.Height) div 2 ;
end;

procedure TFusergrp.RempliJAL ;
var Q : TQuery ;
BEGIN
TJAL:=TStringList.Create ;
Q:=OpenSQL('Select J_JOURNAL,J_LIBELLE FROM JOURNAL ORDER BY J_LIBELLE',True) ;
While Not Q.EOF do
  BEGIN
  TJAL.Add(Q.Fields[0].AsString) ;
  JOURNALAUTORISE.Items.Add(Q.Fields[1].AsString+' ('+Q.Fields[0].AsString+')') ;
  Q.Next ;
  END ;
Ferme(Q) ;
MajSelect ;
END ;

procedure  TFusergrp.MajSelect ;
var ListeJal : String;
    Code : String ;
    Ok : boolean ;
    i : integer ;
BEGIN
if UG_JALAUTORISES.Text='' then Ok:=True else Ok:=False ;
for i:=0 to JOURNALAUTORISE.Items.Count-1 do JOURNALAUTORISE.Selected[i]:=Ok ;
if Ok then Exit ;
ListeJal:=UG_JALAUTORISES.Text ;
While Length(ListeJal)<>0 do
  BEGIN
  Code:=ReadTokenSt(ListeJal) ;
  if TJAL.IndexOf(Code) <>-1 then JOURNALAUTORISE.Selected[TJAL.IndexOf(Code)]:=True ;
  END
END ;

procedure TFusergrp.BTagClick(Sender: TObject);
begin
  inherited;
TagDetag(TRUE) ;
end;

procedure TFusergrp.BdetagClick(Sender: TObject);
begin
  inherited;
TagDetag(FALSE) ;
end;

procedure TFUserGrp.TagDetag (Ok : Boolean) ;
var i : integer ;
BEGIN
BTag.Visible:=not Ok ;
BDeTag.Visible:=Ok ;
for i:=0 to JOURNALAUTORISE.Items.Count - 1 do
  if JOURNALAUTORISE.Selected[i]=not Ok then JOURNALAUTORISE.Selected[i]:=Ok ;
END ;

procedure TFusergrp.UG_JALAUTORISESExit(Sender: TObject);
var St : String ;
begin
  inherited;
St:=UG_JALAUTORISES.Text ;
if St='' then Exit ;
if St[Length(St)]<>';' then UG_JALAUTORISES.Text:=St+';' ;
end;

Function TFUsergrp.VerifCodeJal : boolean ;
Var Code,St : string ;
BEGIN
Result:=False ;
St:=UG_JALAUTORISES.Text ;
While length(St)<>0 do
  BEGIN
  Code:=ReadTokenSt(St) ;
  if TJAL.IndexOf(Code)=-1 then BEGIN HM2.execute(0,'','') ; UG_JALAUTORISES.SetFocus ; Exit ; END ;
  END;
Result:=True ;
END ;

Procedure TFUsergrp.ChargMagUserGrp ;
BEGIN
if V_PGI.UserGrp=TaUG_NUMERO.AsInteger then
   BEGIN
   V_PGI.UserGrp:=TaUG_NUMERO.AsInteger ;
   V_PGI.Confidentiel:=TaUG_CONFIDENTIEL.AsString ;
   VH^.GrpMontantMin:=TaUG_MONTANTMIN.AsFloat ;
   VH^.GrpMontantMax:=TaUG_MONTANTMAX.AsFloat ;
   VH^.JalAutorises:=TaUG_JALAUTORISES.AsString ;
   if VH^.JalAutorises<>'' then
      BEGIN
      if VH^.JalAutorises[1]<>';' then VH^.JalAutorises:=';'+VH^.JalAutorises ;
      if VH^.JalAutorises[Length(VH^.JalAutorises)]<>';' then VH^.JalAutorises:=VH^.JalAutorises+';' ;
      END ;
   V_PGI.LanguePrinc:=TaUG_LANGUE.AsString ;
   V_PGI.LanguePerso:=TaUG_PERSO.AsString ;
   if V_PGI.LanguePrinc = '' then V_PGI.LanguePrinc := 'FRA' ;
   if V_PGI.LanguePerso = '' then V_PGI.LanguePerso := V_PGI.LanguePrinc ;
   END ;
END ;

procedure TFusergrp.FormClose(Sender: TObject; var Action: TCloseAction);
begin
inherited;
TJAL.Clear ; TJAL.Free ; TJAL:=Nil ;
if Modifier then BEGIN Logout ; LoginUsers ; END ;
end;

procedure TFusergrp.UG_LANGUEClick(Sender: TObject);
var Qm : TQuery ;
//    sp : String ;
begin
//sp := UG_PERSO.Value ;
Qm := OpenSQL('SELECT CO_CODE,CO_LIBELLE FROM COMMUN WHERE CO_TYPE="TRA" AND CO_LIBRE="' + UG_LANGUE.Value + '"',True) ;
UG_PERSO.Items.Clear ; UG_PERSO.Values.Clear ;
UG_Perso.Items.Add(HM2.Mess[1]) ; UG_Perso.Values.Add('') ;
while not Qm.EOF do
   begin
   UG_PERSO.Items.Add(Qm.FindField('CO_LIBELLE').AsString) ;
   UG_PERSO.Values.Add(Qm.FindField('CO_CODE').AsString) ;
   Qm.Next ;
   end ;
Ferme(Qm) ;
//UG_PERSO.Value := sp ;
end;

procedure TFusergrp.BDeleteClick(Sender: TObject);
Var Q : TQuery ;
    Trouv : boolean ;
begin
if V_PGI.Groupe=TaUG_GROUPE.AsString then BEGIN HM2.Execute(3,'','') ; Exit ; END ;
Trouv:=False ;
Q:=OpenSQL('Select Count(US_UTILISATEUR) from UTILISAT Where US_GROUPE="'+TaUG_GROUPE.AsString+'"',True) ;
if Not Q.EOF then if Q.Fields[0].AsInteger>0 then Trouv:=True ;
Ferme(Q) ;
if Trouv then BEGIN HM2.Execute(2,'','') ; Exit ; END ;
  inherited;
end;

procedure TFusergrp.UG_MONTANTMINExit(Sender: TObject);
begin
  inherited;
if UG_MONTANTMIN.Text<>'' then
   BEGIN
   TaUG_MONTANTMIN.DisplayFormat:=StrfMask(V_PGI.OkDecV,'',True) ;
   END ;
end;

procedure TFusergrp.UG_MONTANTMINEnter(Sender: TObject);
begin
  inherited;
if UG_MONTANTMIN.Text<>'' then TaUG_MONTANTMIN.DisplayFormat:='#########.00' ;
end;

procedure TFusergrp.UG_MONTANTMAXEnter(Sender: TObject);
begin
  inherited;
if UG_MONTANTMAX.Text<>'' then TaUG_MONTANTMAX.DisplayFormat:='#########.00' ;
end;

procedure TFusergrp.UG_MONTANTMAXExit(Sender: TObject);
begin
  inherited;
if UG_MONTANTMAX.Text<>'' then TaUG_MONTANTMAX.DisplayFormat:=StrfMask(V_PGI.OkDecV,'',True) ;
end;

procedure TFusergrp.STaDataChange(Sender: TObject; Field: TField);
begin
if V_PGI.Superviseur then BAcces.Enabled:=(Not(Ta.State in[dsEdit,dsInsert]))
                  else BEGIN Pappli.Enabled:=False ; BAcces.Enabled:=False ; END ;
  inherited;
end;

end.
