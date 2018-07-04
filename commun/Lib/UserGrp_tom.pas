{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 30/08/2001
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : USERGRP (USERGRP)
Mots clefs ... : TOM;USERGRP
*****************************************************************}
Unit USERGRP_TOM ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
     eFiche,eFichList, Maineagl,
{$ELSE}
     db,
     {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
     Fiche, FichList,  Fe_Main,
{$ENDIF}
     forms, sysutils,ComCtrls,HCtrls, HEnt1, HMsgBox, UTOM,
{$IFDEF COMPTA}
     Ent1,
{$ENDIF}
     UTob,
     HTB97,
     HPanel ;

const
  msgAucun : string = 'Aucun journal sélectionné';
  msgUnSeul : string = '1 journal sélectionné';
  msgPlusieurs : string = 'journaux sélectionnés';

Type
  TOM_USERGRP = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    private    { Déclarations privées }
    modifJournaux : boolean ;
    NextNumGrp  : Integer ;
    MemoLgu : String ;
    NbSelect : integer;
//    TJAL : TStringList ;
    BACCES : TButton;
    BINSERT : TButton;
    UG_PERSO : THValComboBox;
    UG_LANGUE : THValComboBox;
//    LJAL : TListBox;
//    CBJAL : TCheckbox;
    MVCBJAL : THMultiValComboBox ;
    CBCONFIDENTIEL : TCheckBox ;
    bCharge : boolean;
    bForceChecked : boolean;
    setOfModule : array of integer;//LM20070912
    procedure initSetOfModule (modules:string) ;//LM20070912
    procedure AutoriseNouveau;
    procedure BAccesClick(Sender: TObject);
    procedure TrouveTrou;
    procedure UG_MONTANTMAXEnter(Sender: TObject);
    procedure UG_MONTANTMINEnter(Sender: TObject);
    function VerifCodeJal: boolean;
    procedure ChargMagUserGrp;
    procedure UG_LANGUEClick(Sender: TObject);
//    procedure OnCBJalClick(sender: tobject);
    procedure CBConfidentielOnClick(sender : TObject);
    procedure MajSelect;
//    procedure RempliJAL;
//    procedure OnLJalClick(sender: tobject);
//    procedure UpdateNbSelect;
    procedure UG_MONTANTMAXExit(Sender: TObject);
    procedure UG_MONTANTMINExit(Sender: TObject);
    end ;

Procedure FicheUserGrp (arguments:string='') ;
function ExistDansMulti(val : string ; multi : THMultiValComboBox) : boolean ;

Implementation

uses UserAcc,
{$IFNDEF EAGLCLIENT}
     MajTable,
{$ENDIF}
     HSysMenu,HDB,
     uSATUtil
     ;

{ --------------------------------------------------------------- }

Procedure FicheUserGrp (arguments:string='');
begin
  AGLLanceFiche('YY','YYUSERGROUP','','',arguments) ; //LM20070912
end;

{ --------------------------------------------------------------- }

function RecupArgument(cle, arg: string): string; //LM20070912
var p: integer;
    s: string;
begin
  p := pos(cle, arg);
  if p > 0 then
  begin
    s := trim(copy(arg, p + length(cle), 1000));
    s := trim(gtfs(s, ';', 1));
    p := pos('=', s);
    if p > 0 then s := trim(copy(s, p + 1, 1000));
  end;
  result := s;
end;

procedure TOM_USERGRP.initSetOfModule (modules:string) ;//LM20070912
var s : string ;
    n:integer;
begin
  modules := Trim (modules) ;
  if modules='' then exit ;
  if copy(modules, length(modules)-1, 1)<> ',' then modules:=modules + ',' ;
  n:=nbCarInstring(modules, ',') ;
  setlength(setOfModule, n) ;
  s:='*' ; n:=0 ;
  while s<>'' do
  begin
    s:=readTokenstV(modules) ;
    if isNumeric(s) then
    begin
      setOfModule[n]:= valeurI(s);
      inc(n) ;
    end ;
  end ;
end ;

procedure TOM_USERGRP.OnArgument ( S: String ) ;
var arg : string ;
begin
Inherited ;
    s:=trim(s) ;
    if s<>'' then //LM20070912 : on peut passer en argument les modules à afficher et la possibilité de masquer la "compta" :
    begin         //             MODULES=1,2,3,4;SANSCOMPTA=X;
      arg := recupArgument ('MODULES', s) ;
      initSetOfModule(arg);

      arg := recupArgument ('SANSCOMPTA',s) ;
      if arg='X' then
      begin
        setControlVisible('GBMONTANT', false) ;
        setControlVisible('PJOURNAL', false) ;
      end ;
    end ;

    bForceChecked := false; bCharge := false;
    modifJournaux := false ;
    //TJAL:=TStringList.Create ;
    BINSERT := TButton(GetControl('BINSERT'));
    BACCES := TButton(GetControl('BACCES'));
    BACCES.OnClick := bAccesClick;
    UG_PERSO := THValComboBox(GetControl('UG_PERSO'));
    UG_LANGUE := THValComboBox(GetControl('UG_LANGUE'));
    UG_LANGUE.OnClick := UG_LANGUEClick;
    MVCBJAL := THMultiValComboBox(GetControl('UG_JALAUTORISES'));
    CBCONFIDENTIEL := TCheckBox(GetControl('CBConfidentiel')) ;
    CBCONFIDENTIEL.OnClick := CBConfidentielOnClick;


    //LJAL := TListBox(Getcontrol('LJAL'));
    //LJAL.OnClick := OnLJalClick;
    //CBJAL := TCheckbox(Getcontrol('CBJAL'));
    //CBJAL.OnClick := OnCBJalClick;
    NbSelect := 0;
    //SetControlText('NBJAL',TraduireMemoire('Aucun journal sélectionné'));
    MemoLgu:='' ;

    if (EstSerie(S3) or (EstSerie(S5) {$IFDEF COMPTA} and not VH^.OkModMultilingue{$ENDIF})) then
    begin
//      SetControlProperty('TUG_Langue','Visible',False); SetControlProperty('UG_Langue','Visible',False);
//      SetControlProperty('TUG_Perso','Visible',False); SetControlProperty('UG_Perso','Visible',False);
    {$IFNDEF GCGC}
      if EstSerie(S3) then
      BEGIN
        SetControlProperty('CBCONFIDENTIEL','Visible',False) ;
        SetControlProperty('UG_NIVEAUACCES','Visible',False) ;
        SetControlProperty('TUG_NIVEAUACCES','Visible',False) ;
      END ;
    {$ENDIF !GCGC}
    END ;
 
  // C.B 21/04/2006
  // déblocage de la sélection par langue pour tests en attendant les modifs annoncées de Joel Sidos
	if (ctxMode in V_PGI.PGIContexte) or (ctxAffaire in V_PGI.PGIContexte) then
  BEGIN
	  SetControlProperty('TUG_LANGUE','Visible',True); SetControlProperty('UG_Langue','Visible',True);
  	SetControlProperty('TUG_PERSO','Visible',True); SetControlProperty('UG_Perso','Visible',True);
  END ;

	if (ctxGescom in V_PGI.PGIContexte) or (ctxAffaire in V_PGI.PGIContexte) or  (V_PGI.LaSerie < S5)  // or ctxPaie in V_PGI.PGIContexte
   then
   BEGIN
   //SetControlVisible('UG_JALAUTORISES',False) ; SetControlVisible('TUG_JALAUTORISES',False) ;
   SetControlVisible('PJOURNAL', False);
   SetControlVisible('GBMontant',False) ;
   //SteControlVisible('ZoomJAL',False) ;
   END ;
    //RR RempliJAL ;
SetControlProperty('UG_MONTANTMIN','DisplayFormat',StrfMask(V_PGI.OkDecV,'',True)) ;
SetControlProperty('UG_MONTANTMAX','DisplayFormat',StrfMask(V_PGI.OkDecV,'',True)) ;
AutoriseNouveau ;
if V_PGI.Superviseur<>True then FicheReadOnly(Ecran);
//else if(Ta.Eof) And (Ta.Bof) And (FTypeAction<>taConsult)then BinsertClick(Nil) ;
TPageControl(GetControl('P')).ActivePage := TTabSheet(GetControl('PGENERAL'));

THDBEdit(GetControl('UG_MONTANTMIN')).OnEnter := UG_MONTANTMINEnter;
THDBEdit(GetControl('UG_MONTANTMIN')).OnExit := UG_MONTANTMINExit;
THDBEdit(GetControl('UG_MONTANTMAX')).OnEnter := UG_MONTANTMAXEnter;
THDBEdit(GetControl('UG_MONTANTMAX')).OnExit := UG_MONTANTMAXExit;

//uniquement en line
//TToolBarButton97(GetCOntrol('BACCES')).Visible := false;

end ;

procedure TOM_USERGRP.OnLoadRecord ;
begin
  Inherited ;
if V_PGI.Groupe=GetField('UG_GROUPE') then SetControlProperty('UG_PASSWORD','PasswordChar',#0)
                                      else SetControlProperty('UG_PASSWORD','PasswordChar','*');
if V_PGI.Superviseur<>True then FicheReadOnly(Ecran) ;
if MemoLgu<>GetField('UG_LANGUE') then UG_LANGUEClick(nil) ;
if GetField('UG_PERSO')='' then UG_PERSO.ItemIndex := UG_PERSO.Items.IndexOf(TraduireMemoire('Aucune'))
                           else UG_PERSO.ItemIndex := UG_PERSO.Values.Indexof(GetField('UG_PERSO')) ;
MemoLgu:=GetField('UG_LANGUE');
MajSelect;
end ;

procedure TOM_USERGRP.OnNewRecord ;
begin
  Inherited ;
BAcces.Enabled:=False ;
TrouveTrou;
SetField('UG_NUMERO',NextNumGrp) ;
end ;

procedure TOM_USERGRP.OnDeleteRecord ;
Var i,j : Byte ;
    joker : String ;
    StLike : String[100] ;
    StLike1 : String[100] ;
    function TrouveUsers : boolean;
    begin
    LastErrorMsg:='';
    if V_PGI.Groupe=GetField('UG_GROUPE') then LastErrorMsg:='Vous ne pouvez pas supprimer votre propre groupe';
    result := ExisteSQL('Select US_UTILISATEUR from UTILISAT Where US_GROUPE="'+GetField('UG_GROUPE')+'"');
    if result then LastErrorMsg:='Ce groupe possède des utilisateurs. La suppression est impossible.';
    if LastErrorMsg<>'' then LastError:=1 else LastError:=0;

{
    if V_PGI.Groupe=GetField('UG_GROUPE') then BEGIN PGIInfo('Vous ne pouvez pas supprimer votre propre groupe','Groupes d''utilisateurs') ; Exit ; END ;
    result := ExisteSQL('Select Count(US_UTILISATEUR) from UTILISAT Where US_GROUPE="'+GetField('UG_GROUPE')+'"');
    if result then PGIInfo('Ce groupe possède des utilisateurs. La suppression est impossible.','Groupes d''utilisateurs');
}
    end;
begin
if TrouveUsers then exit;
Inherited;
i:=GetField('UG_NUMERO') ; StLike:='' ; StLike1:='' ;
for j:=1 to i-1 do StLike:=StLike+'_' ;
StLike1:=StLike ; StLike:=StLike+'X' ; StLike1:=StLike1+'-' ;
if V_PGI.Driver=dbMSACCESS then Joker:='*' else Joker:='%' ;
ExecuteSQL('UPDATE MENU SET MN_ACCESGRP=(substring(MN_ACCESGRP,1,' + intToStr(i-1) + ')' +
              '||"0"||substring(MN_ACCESGRP,' + intToStr(i+1) + ','+intToStr(100-i) + ')) ' +
              'WHERE MN_ACCESGRP LIKE "' + StLike + Joker + '" OR MN_ACCESGRP LIKE "' + StLike1 + Joker + '"') ;
ExecuteSql('Delete From UTILISAT Where US_GROUPE="'+GetField('UG_GROUPE')+'"') ;
AutoriseNouveau ;
end ;

procedure TOM_USERGRP.OnUpdateRecord ;
{var Tous : boolean;
    St,StJal,StMess : string;
    i : integer; }
begin
  if MVCBJAL.Value='' then MVCBJAL.Value:='-' ;
{
    tous := false ;
    if MVCBJAL.Tous=true then tous:=true ;

    if Tous then St:='' else
      begin
      for i:=0 to MVCBJAL.Items.Count-1 do
        BEGIN
          St:=St+MVCBJAL.Items[i]+';';
        END ;
      //RRO11/12/2002
      //Correctionde l'erreur de sauvegarde des sélections manuel de journal
      StJal := st ;
      end;
    if Length(St)>200 then
      begin
      StJal := '';
      StMess := 'Sélection partielle maximum.'#10#13+
                'La liste va être troquée.'#10#13+'Souhaitez poursuivre la validation ?';
      if PGIAsk(StMess,'Sélection des journaux autorisés')=mrYes then
        begin
        while (St<>'') and (Length(StJal)<=196) do
          StJal := StJal + ReadTokenSt(St) + ';';
        end
        else
        begin
        LastErrorMsg:='Enregistrement annulé';
        LastError:=1;
        exit;
        end;
      end;
    SetField('UG_JALAUTORISES',StJal);
}    
    Inherited ;
end ;

procedure TOM_USERGRP.OnAfterUpdateRecord ;
begin
  if MVCBJAL.Value='-' then MVCBJAL.Value:='';
{$IFNDEF EAGLCLIENT}
TFFicheListe(Ecran).Modifier:=True ;
{$ENDIF}
if DS.state in [dsEdit,dsInsert] then
   BEGIN
//   if TaUG_ABREGE.AsString='' then BEGIN HM.Execute(6,'','') ; UG_ABREGE.SetFocus ; Exit ; END ;
//   if TaUG_PASSWORD.AsString='' then BEGIN HM.Execute(7,'','') ; UG_PASSWORD.SetFocus ; Exit ; END ;
   if not VerifCodeJal then Exit ;
   END ;
{$IFNDEF EAGLCLIENT}
TFFicheListe(Ecran).Modifier:=False ;
{$ENDIF}
ChargMagUserGrp ;
Inherited ;
//AutoriseNouveau ;
end ;

procedure TOM_USERGRP.OnChangeField ( F: TField ) ;
//var St : String ;
begin
Inherited ;
{if F.FieldName='UG_JALAUTORISES' then
  begin
  St:=GetField('UG_JALAUTORISES');
  if (St<>'') and (St[Length(St)]<>';') then SetField('UG_JALAUTORISES',St+';');
  end else
}
if F.FieldName='UG_MONTANTMIN' then
  begin
  if GetField('UG_MONTANTMIN')<>0 then
    SetControlProperty('UG_MONTANTMIN','DisplayFormat',StrfMask(V_PGI.OkDecV,'',True)) ;
  end else
if F.FieldName='UG_MONTANTMAX' then
  begin
  if GetField('UG_MONTANTMAX')<>0 then
    SetControlProperty('UG_MONTANTMAX','DisplayFormat',StrfMask(V_PGI.OkDecV,'',True)) ;
  end else

BAcces.Enabled:=(Not(DS.State in[dsEdit,dsInsert])) and (V_PGI.Superviseur);
end ;

procedure TOM_USERGRP.OnClose ;
begin
  Inherited ;
//RR TJAL.Clear ; TJAL.Free ; TJAL:=Nil ;
{$IFNDEF EAGLCLIENT}
if TFFicheListe(Ecran).Modifier then BEGIN Logout ; LoginUsers ; END ;
{$ENDIF}
end ;

procedure TOM_USERGRP.OnCancelRecord ;
begin
  Inherited ;
MajSelect ;
end ;

{ --------------------------------------------------------------- }


procedure TOM_USERGRP.UG_MONTANTMINEnter(Sender: TObject);
begin
  inherited;
if GetField('UG_MONTANTMIN')<>0 then SetControlProperty('UG_MONTANTMIN','DisplayFormat','#########.00') ;
end;

procedure TOM_USERGRP.UG_MONTANTMAXEnter(Sender: TObject);
begin
  inherited;
if GetField('UG_MONTANTMAX')<>0 then SetControlProperty('UG_MONTANTMAX','DisplayFormat','#########.00') ;
end;

procedure TOM_USERGRP.UG_MONTANTMINExit(Sender: TObject);
begin
  inherited;
if GetField('UG_MONTANTMIN')<>0 then SetControlProperty('UG_MONTANTMIN','DisplayFormat',StrfMask(V_PGI.OkDecV,'',True)) ;
end;

procedure TOM_USERGRP.UG_MONTANTMAXExit(Sender: TObject);
begin
  inherited;
if GetField('UG_MONTANTMAX')<>0 then SetControlProperty('UG_MONTANTMAX','DisplayFormat',StrfMask(V_PGI.OkDecV,'',True)) ;
end;


procedure TOM_USERGRP.BAccesClick(Sender: TObject);
Var
  SS:SetInteger ;

const SSctxGRC       :array[0..0]  of integer = (92);
const SSctxCHRHS3    :array[0..7]  of integer = (161,162,92,165,166,26,112,107);
const SSctxCHRPMS5   :array[0..9]  of integer = (161,162,288,92,165,166,26,112,107,160);
const SSctxRMSS3     :array[0..7]  of integer = (286,162,92,165,166,26,112,111);
const SSctxRMSS5     :array[0..10]  of integer = (161,286,162,288,92,165,166,26,112,160,111);
const SSctxBTP       :array[0..5]  of integer = (145,149,60,148,26,27);
const SSctxScot      :array[0..11]  of integer = (141,142,143,144,151,152,154,26,27,292,293,294);
const SSctxAffaire   :array[0..12] of integer = (71,72,73,138,139,92,74,140,153,26,27,262,160);
const SSctxFO        :array[0..6]  of integer = (107,108,109,111,112,26,27);
const SSctxMode      :array[0..17] of integer = (101,102,103,59,104,105,106,107,108,109,110,111,112,113,114,115,26,27);
{$IFDEF GESCOM}
const SSctxGescom    :array[0..14] of integer = (30,36,31,32,33,70,65,60,27,26,59,160,111,267,260);
{$ELSE}
const SSctxGescom    :array[0..9]  of integer = (30,31,32,33,70,65,60,27,26,59);
{$ENDIF}
const SSctxJuridique :array[0..3]  of integer = (85,87,187,261);
const SSctxPaie      :array[0..6]  of integer = (41,42,43,44,46,47,49);
// GP
const SSctxImmo      :array[0..14]  of integer = (250,251,252,253,254,255,256,258,259,183,184,185,26,312,313);
const SSctxPCL       :array[0..7]  of integer = (52,53,54,55,56,96,26,27);
const SSCCMPS3       :array[0..4]  of integer = (37,38,17,18,25);    // Modif Fiche 12088 SBO
const SSCCMP         :array[0..5]  of integer = (37,38,39,17,18,25); // Modif Fiche 12088 SBO
const SSComptaS3     :array[0..11] of integer = (6,8,9,11,13,14,16,17,18,20,26,27);
const SSComptaS5     :array[0..12] of integer = (6,8,9,11,12,13,14,16,17,18,20,26,27);
//const SSComptaS7     :array[0..8]  of integer = (1,2,3,6,8,7,15,25,26);
const SSComptaS7     :array[0..12]  of integer = (6,8,9,11,12,13,14,16,17,18,20,26,27);
const SSCtxDP        :array[0..21] of integer = (75,76,78,181,182,85,87,261,41,42,43,44,46,47,49,141,142,143,144,151,152,154);
const SSCtxDP_PCL    :array[0..7]  of integer = (52,53,54,55,56,96,26,27);
const SSCtxDP_S5     :array[0..7]  of integer = (1,2,3,6,7,15,25,26);
const SSCtxDP_S7     :array[0..12] of integer = (6,8,9,11,12,13,14,16,17,18,20,26,27); // Modif PACK AVANCE SBO
const SSctxGPAO      :array[0..17] of integer = (120,121,122,125,126,210,211,212,213,215,219,26,27,260,70,59,36,160);
const SSctxTRESO     :array[0..7]  of integer = (130,131,132,133,134,135,136,27);

      Procedure  InitModuleAcces ( const AA : array of integer  );
      var i  :  integer;
      begin
        SetLength(SS,High(AA)+1);
        For i:=0 to High(AA) do SS[i]:=AA[i];
      end;
      Procedure  MergeModuleAcces ( const AA : array of integer  );
      var i  :  integer;
      begin
        For i:=0 to High(AA) do
        begin
           SetLength(SS,Length(SS)+1);
           SS[High(SS)]:=AA[i];
        end;
      end;

begin
inherited;
 if high(setofModule)>-1 then //LM20070912
    InitModuleAcces(setofModule)
 else if ctxBTP in V_PGI.PGIContexte then InitModuleAcces(SSctxBTP) else
 if ctxAffaire in V_PGI.PGIContexte then
 BEGIN
   if ctxScot in V_PGI.PGIContexte then InitModuleAcces(SSctxScot)
                                   else InitModuleAcces(SSctxAffaire);
 END else
  if ctxMode in V_PGI.PGIContexte then
  BEGIN
    if ctxFO in V_PGI.PGIContexte then InitModuleAcces(SSctxFO)
    else InitModuleAcces(SSctxMode);
  END else
   if ctxGPAO in V_PGI.PGIContexte then InitModuleAcces(SSctxGPAO) else
    if ctxGescom in V_PGI.PGIContexte then InitModuleAcces(SSctxGescom) else
     if ctxJuridique in V_PGI.PGIContexte then InitModuleAcces(SSctxJuridique) else
      if ctxPaie in V_PGI.PGIContexte then InitModuleAcces(SSctxPaie) else
// GP
{$IFDEF PGIIMMO}
      If ctxImmo in V_PGI.PGIContexte then InitModuleAcces(SSctxImmo) Else
{$ENDIF}
      {$IFDEF TRESO}
       if ctxTreso in V_PGI.PGIContexte then InitModuleAcces(SSctxTRESO) else
      {$ENDIF}
        BEGIN //compta par defaut
         if ctxPCL in V_PGI.PGIContexte then InitModuleAcces(SSctxPCL) else
         begin
{$IFDEF CCMP}
           // Modif Fiche 12088
           if EstSerie(S3)
             then InitModuleAcces(SSCCMPS3)
             else InitModuleAcces(SSCCMP) ;
           // Fin Modif Fiche 12088
{$ELSE}
           if EstSerie(S7) then InitModuleAcces(SSComptaS7) else
            if EstSerie(S5) then  InitModuleAcces(SSComptaS5) else
             if EstSerie(S3) then  InitModuleAcces(SSComptaS3) ;
{$ENDIF}
         end;
        END;

{$IFNDEF MODE}
if ctxGRC in V_PGI.PGIContexte then
BEGIN
   MergeModuleAcces(SSctxGRC);  // attention merge pour ajouter module GRC
END;
{$ENDIF}

if ctxCHR in V_PGI.PGIContexte then
begin
{$IFDEF RMSCHR}
  if EstSerie(S5) then InitModuleAcces(SSCtxRMSS5);
  if EstSerie(S3) then InitModuleAcces(SSCtxRMSS3);
{$ELSE}
  if EstSerie(S5) then InitModuleAcces(SSCtxCHRPMS5);
  if EstSerie(S3) then InitModuleAcces(SSCtxCHRHS3);
{$ENDIF}
end;

// lanceur = tous les menus des applications multi-dossiers
// ok si AGL >= v504c
if ctxDP in V_PGI.PGIContexte then
BEGIN
   InitModuleAcces(SSCtxDP); // DP
   // compta
   if ctxPCL in V_PGI.PGIContexte then MergeModuleAcces(SSCtxDP_PCL) else
   begin
     if EstSerie(S7) then MergeModuleAcces(SSCtxDP_S7) else
      if EstSerie(S5) then MergeModuleAcces(SSCtxDP_S5) ;
   end;
END ;

FicheUserAcces(GetField('UG_GROUPE'),GetField('UG_LIBELLE'),GetField('UG_NUMERO'),SS,'',False) ; // PCS 11/10/2004 suppression de blocage seul utilisateur

SS:=Nil;

end;

procedure TOM_USERGRP.UG_LANGUEClick(Sender: TObject);
var Qm : TQuery ;
begin
Qm := OpenSQL('SELECT CO_CODE,CO_LIBELLE FROM COMMUN WHERE CO_TYPE="TRA" AND CO_LIBRE="' + UG_LANGUE.Value + '"',True,-1,'',true) ;
UG_PERSO.Items.Clear ; UG_PERSO.Values.Clear ;
UG_Perso.Items.Add(TraduireMemoire('Aucune')) ; UG_Perso.Values.Add('') ;
while not Qm.EOF do
   begin
   UG_PERSO.Items.Add(Qm.FindField('CO_LIBELLE').AsString) ;
   UG_PERSO.Values.Add(Qm.FindField('CO_CODE').AsString) ;
   Qm.Next ;
   end ;
Ferme(Qm) ;
end;

procedure TOM_USERGRP.CBConfidentielOnClick(sender : TObject);
begin
    if not bCharge then
    begin
        DS.Edit;
        if CBCONFIDENTIEL.Checked then SetField('UG_CONFIDENTIEL','1')
        else SetField('UG_CONFIDENTIEL','0') ;
    end;
end;

{
procedure TOM_USERGRP.OnCBJalClick(sender : tobject);
var i : integer;
begin
if not bForceChecked then
  for i:=0 to LJAL.Items.Count-1 do
    LJAL.Selected[i]:=CBJAL.Checked;
if not bCharge then
  if (not (DS.state in [dsEdit,dsInsert])) then ForceUpdate;
UpdateNbSelect;
end;
}

(*
procedure TOM_USERGRP.OnLJalClick(sender : tobject);
begin
{$IFNDEF EAGL}
if not bCharge then
  if (not (DS.state in [dsEdit,dsInsert])) then ForceUpdate;
{$ENDIF}
UpdateNbSelect;
bForceChecked := true;
CBJAL.Checked := (LJAL.SelCount=LJAL.Items.Count);
bForceChecked := false;
end;
*)
{ --------------------------------------------------------------- }

(*
Procedure TOM_USERGRP.UpdateNbSelect;
var msg : string;
begin
//nbSelect := LJAL.SelCount;
if nbSelect = 0 then msg := msgAucun
else if nbSelect = 1 then msg := msgUnSeul else
msg := IntToStr(nbSelect)+' '+msgPlusieurs;
SetControlText('NBJAL',TraduireMemoire(msg));
end;
*)

Procedure TOM_USERGRP.TrouveTrou  ;
Var Q : TQuery ;
BEGIN
(*
Q:=OpenSQL('Select UG_NUMERO from USERGRP order by UG_NUMERO',TRUE,-1,'',true) ;
While(Not Q.EOF)AND(Q.Fields[0].AsInteger=NextNumGrp) do BEGIN  Q.Next ; Inc(NextNumGrp) ;  END ;
if(Q.EOF)AND(NextNumGrp>100) then NextNumGrp:=0 ;
Ferme(Q) ;
*)
Q:=OpenSQL('SELECT MAX(UG_NUMERO) AS MAX FROM USERGRP',TRUE,-1,'',true) ;
if not Q.eof then
begin
  NextNumGrp:=Q.fields[0].asInteger + 1 ;
end else
begin
  NextNumGrp:=1 ;
end;
ferme (Q);
END ;

Procedure TOM_USERGRP.AutoriseNouveau ;
BEGIN
TrouveTrou ; BInsert.Enabled:=(NextNumGrp>0) ; BAcces.Enabled:=(V_PGI.Superviseur=True) ;
END ;

function ExistDansMulti(val : string ; multi : THMultiValComboBox) : boolean ;
var i : integer ;
begin
  i := 0 ;
  result := false ;
  while (i < Multi.Items.Count - 1) do
  begin
    if val=Multi.Values[i] then
    begin
      result := true ;
      exit;
    end ;
    i:= i+1 ;
  end ;
end ;

Function TOM_USERGRP.VerifCodeJal : boolean ;
Var Code,St, stBoucle, stVal{,msg} : string ;
//    valeur : string ;
    erreur : boolean ;
BEGIN
Result:=False ;
St:=GetField('UG_JALAUTORISES') ;
stBoucle := St ;
erreur := false ;
NbSelect:=0 ;
	if not MVCBJAL.Tous then
    while Length(StBoucle)>0 do
   	begin
	    stVal := READTOKENST(stBoucle);
      //if not(ExistDansMulti(stVal,MVCBJAL)) then erreur := true ;
      //if MVCBJAL.items.IndexOf(stVal)= -1 then erreur := true;
    end;

if erreur then
begin
  PGIInfo('Un des codes journaux que vous avez saisi n''existe pas.','Groupes d''utilisateurs');
  SetFocusControl('UG_JALAUTORISES') ;
end ;

While length(St)<>0 do
  BEGIN
  Code:=ReadTokenSt(St) ;
  //if TJAL.IndexOf(Code)=-1 then
  if MVCBJAL.Items.IndexOfName(Code) =-1 then
    BEGIN
    PGIInfo('Un des codes journaux que vous avez saisi n''existe pas.','Groupes d''utilisateurs');
    SetFocusControl('UG_JALAUTORISES') ;
    Exit ;
    END ;
  END;

Result:=True ;
END ;

Procedure TOM_USERGRP.ChargMagUserGrp ;
BEGIN
if V_PGI.UserGrp=GetField('UG_NUMERO') then
   BEGIN
//   V_PGI.UserGrp:=GetField('TaUG_NUMERO') ;
   V_PGI.Confidentiel:=GetField('UG_CONFIDENTIEL');
//EPZ A REMETTRE AVEC USES ENT1....
//EPZ    VH^.GrpMontantMin:=GetField('UG_MONTANTMIN') ;
//EPZ   VH^.GrpMontantMax:=GetField('UG_MONTANTMAX') ;
//EPZ   VH^.JalAutorises:=GetField('UG_JALAUTORISES') ;
//EPZ   if VH^.JalAutorises<>'' then
      BEGIN
//EPZ      if VH^.JalAutorises[1]<>';' then VH^.JalAutorises:=';'+VH^.JalAutorises ;
//EPZ      if VH^.JalAutorises[Length(VH^.JalAutorises)]<>';' then VH^.JalAutorises:=VH^.JalAutorises+';' ;
      END ;
   V_PGI.LanguePrinc:=GetField('UG_LANGUE');
   V_PGI.LanguePerso:=GetField('UG_PERSO');
   if V_PGI.LanguePrinc = '' then V_PGI.LanguePrinc := 'FRA' ;
   if V_PGI.LanguePerso = '' then V_PGI.LanguePerso := V_PGI.LanguePrinc ;
   END ;
END ;

{ --------------------------------------------------------------- }
(*
procedure TOM_USERGRP.RempliJAL ;
//var Q : TQuery ;
BEGIN
{Q:=OpenSQL('Select J_JOURNAL,J_LIBELLE FROM JOURNAL ORDER BY J_LIBELLE',True) ;
TJAL.Clear;
While Not Q.EOF do
  BEGIN
  TJAL.Add(Q.FindField('J_JOURNAL').AsString) ;
  LJAL.Items.Add(Q.FindField('J_LIBELLE').AsString+' ('+Q.FindField('J_JOURNAL').AsString+')') ;
  Q.Next ;
  END ;
Ferme(Q) ;
// MajSelect ;
}
END ;
*)

procedure  TOM_USERGRP.MajSelect ;
{var ListeJal : String;
    Code : String ;
    Ok : boolean ;
    i : integer ; }
BEGIN
bCharge:=true;
if getField('UG_CONFIDENTIEL') = '1' then
  CBCONFIDENTIEL.Checked:= true
else
  CBCONFIDENTIEL.Checked := False ;
{
ListeJal:=GetField('UG_JALAUTORISES') ;
if ListeJal='' then Ok:=True else Ok:=False ;
CBJAL.Checked := Ok;
if Ok then OnCBJalClick(nil)
      else
        begin
        for i:=0 to LJAL.Items.Count-1 do LJAL.Selected[i]:=false;
        While Length(ListeJal)<>0 do
          BEGIN
          Code:=ReadTokenSt(ListeJal) ;
          LJAL.Selected[TJAL.IndexOf(Code)]:= (TJAL.IndexOf(Code) <>-1);
          END;
        UpdateNbSelect;
        end;
}
bCharge:=false;
//CBJAL.Checked := (LJAL.SelCount=LJAL.Items.Count);
end;

{
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

}


Initialization
  registerclasses ( [ TOM_USERGRP ] ) ;
end.
