{***********UNITE*************************************************
Auteur  ...... : TJ
Créé le ...... : 15/05/2006
Modifié le ... : 29/05/2006
Description .. : Source TOM de la TABLE : CIBLAGE (CIBLAGE)
Mots clefs ... : TOM;CIBLAGE
*****************************************************************}
Unit UtomCIBLAGE ;

Interface

Uses
 StdCtrls,
{$IFNDEF EAGLCLIENT}
     db,
{$else}
{$ENDIF}
     HMsgBox,
     HCtrls,
     forms,
     sysutils,
     Utob,
     HTB97,
     UTOM;


Type
  TOM_CIBLAGE = Class (TOM)
  public

    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
  private
    Rb2                 : TRadioButton;
    Rb3                 : TRadioButton;
    Anc_Mode            : string;
    CodeCiblage         : String;
    bCiblageAuto        : Boolean;
    Bbloque             : Boolean;
    BAvert              : Boolean;
    procedure BtUserOnClick ( Sender : TObject);
//    procedure ThResponsableOnChange ( Sender : TObject );
    procedure MajCodeCiblage ( Sender : TObject );
    procedure BselectOnClick(Sender: TObject);
    procedure RbOnClick (Sender : Tobject);
    procedure ThResponsableOnElipsisclick (Sender : Tobject);
    end ;

Implementation
uses
  KpmgUtil,    //InitRadioBtModeCiblage
  HEnt1,
  Classes,
  EntRT,
  UtilRessource,
  AglInitGc,
  HDB,
  ParamSoc,
{$IFNDEF EAGLCLIENT}
     Fiche,
{$else}
     eFiche,
{$ENDIF}
  Controls,
  UTofRTTOPCIBLE,
  Windows, //GetKeyState
  UTofRTTopTiers_TOF,
  UTofRtCibleSus_Tof,
  UtofRTTopContact_tof,
  UtofRTTopTiersCon_Tof,
  UtofRTTopSuscon_Tof,
  UtilGc;

const texteMessage: array[1..5] of string = (
    {1}  'Confirmation'
    {2}, 'Cette cible contient des éléments déjà sélectionnés, voulez-vous vraiment la supprimer ?'
    {3}, 'Suppression impossible : ciblage validé'
    {4}, 'Des adresses risquent d''être enlevées du ciblage'
    {5}, 'Le code responsable n''existe pas'
    );


procedure TOM_CIBLAGE.OnArgument ( S: String ) ;
var
  i                     : integer;
  Bselect               : TToolbarButton97;
  Critere               : String;
  StrArg                : String;
  StrVal                : String;
  x                     : integer;
  F                     : TForm;
  ED: tcontrol;

    BtUser              : TToolBarButton97;
 //   ThResponsable       : THEdit;
    ThCodeCiblage       : THEdit;
    Rb1                 : TRadioButton;


begin
  Inherited ;
  Bbloque               := False;
  BAvert                := False;

  Critere               := UpperCase_(Trim(READTOKENST(S)));
  While Critere <> '' do
  begin
    x                   := Pos('=', Critere);
    if x <> 0 then
    begin
      StrArg            := Copy(Critere, 1, x-1);
      StrVal            := Copy(Critere, x+1, length(Critere));
    end
    else
    begin
      StrArg            := Critere;
      StrVal            := '';
    end;

    if StrArg = 'BLOQUE' then
      Bbloque           := StrToBool_(StrVal);

    if StrArg = 'AVERTISSEMENT' then
      BAvert            := StrToBool_(StrVal);

    Critere             := UpperCase_(Trim(READTOKENST(S)));
  end;

  F                     := TForm(Ecran);
  InitRadioBtModeCiblage(F, Bbloque);

  Rb1                   := ThRadioButton(GetControl('RB1'));
  Rb2                   := ThRadioButton(GetControl('RB2'));
  Rb3                   := ThRadioButton(GetControl('RB3'));
  if Assigned(Rb1) and Assigned(Rb2) and Assigned(Rb3) then
  begin
    Rb1.OnClick         := RbOnClick;
    Rb2.OnClick         := RbOnClick;
    Rb3.OnClick         := RbOnClick;
    SetControlText('RCB_MODECIBLAGE', '001');
  end;


  BtUser                := TToolBarButton97(GetControl('BTUSER'));
  If Assigned (BtUser) then BtUser.OnClick        := BtUserOnClick;

  Ed := getcontrol('RCB_RESPONSABLE');
  If Assigned(Ed) then
    if (Ed is THDBEdit) then
    begin
  //ThResponsable.OnChange := ThResponsableOnChange;
      THDBEdit(ed).OnElipsisClick := ThResponsableOnElipsisclick;
    end;

  ThCodeCiblage         := THEdit(GetControl('RCB_CODECIBLAGE'));
  If Assigned (ThCOdeCiblage) then ThCodeCiblage.OnExit  := MajCodeCiblage;

   //Fq 10440
//   SetControlVisible('NOMCOMPLETUSER', False);
   //correction suite présentation kpmg
//   SetControlVisible('NOMCOMPLETUSER', True);

  // maj des champs libres
  for i := 1 to 5 do
  begin
   ChangeLibre2('TRCB_RCBTABLELIBRE'+intToStr(i), Tform(Ecran));
   ChangeLibre2('TRCB_RCBTEXTELIBRE'+intToStr(i), Tform(Ecran));
   ChangeLibre2('TRCB_RCBDATELIBRE'+intToStr(i), Tform(Ecran));
  end;

  Bselect               := TToolBarButton97(getcontrol('BSELECT'));
  if Assigned(Bselect) then
    Bselect.OnClick     := BselectOnClick;

  bCiblageAuto          := False;  
end ;

procedure TOM_CIBLAGE.BselectOnClick(Sender: TObject);
begin
  //verif validation ciblage
  if GetField('RCB_VALIDATION')='X' then
  begin
     PGIBox('Ciblage déjà validé');
     exit;
  end;

  if ((GetField('RCB_MODECIBLAGE')='002') or (GetField('RCB_MODECIBLAGE')='003')) and Bbloque then
    exit;

  if GetField('RCB_CODECIBLAGE') = '' then
     exit;

  //si ciblage sur Raison sociale
  if GetField('RCB_TYPECIBLAGE')='001' then
  begin
     if GetField('RCB_MODECIBLAGE')='003' then //mixte
        RTLanceFiche_RTTOPCIBLE('RT', 'RTTOPCIBLE', '', '', Getfield('RCB_CODECIBLAGE')+';'+Getfield('RCB_MODECIBLAGE'));
     if GetField('RCB_MODECIBLAGE')='001' then //Tiers
        RTLanceFiche_RTTOPTIERS('RT', 'RTTOPTIERS', '', '', Getfield('RCB_CODECIBLAGE')+';'+Getfield('RCB_MODECIBLAGE'));
     if GetField('RCB_MODECIBLAGE')='002' then //Suspect
        RTLanceFiche_RTCIBLESUS('RT', 'RTCIBLESUS', '', '', Getfield('RCB_CODECIBLAGE')+';'+Getfield('RCB_MODECIBLAGE'));
  end
  else
  begin
     if GetField('RCB_MODECIBLAGE') = '003' then  // mixte
        RTLanceFiche_UTOFRTTOPCONTACT('RT', 'RTTOPCONTACT', '', '', Getfield('RCB_CODECIBLAGE')+';'+Getfield('RCB_MODECIBLAGE'));
     if GetField('RCB_MODECIBLAGE') = '001' then // tiers
        RTLanceFiche_RTTOPTIERSCON('RT', 'RTTOPTIERSCON', '', '', Getfield('RCB_CODECIBLAGE')+';'+Getfield('RCB_MODECIBLAGE'));
     if GetField('RCB_MODECIBLAGE') = '002' then // suspect
        RTLanceFiche_RTTOPSUSCON('RT', 'RTTOPSUSCON', '', '', Getfield('RCB_CODECIBLAGE')+';'+Getfield('RCB_MODECIBLAGE'));
  end;

end;


procedure TOM_CIBLAGE.OnNewRecord ;
//var
//   CompteurCiblage      : integer;
begin
  Inherited ;

  SetControlEnabled('RCB_TYPECIBLAGE',true); // mng pour le cas ou l'on fait nouveau sur une fiche existante, il était grisé
  SetField ('RCB_MODECIBLAGE', '001');      // par défaut, "client prospect"
  SetField ('RCB_TYPECIBLAGE', '001');      // par défaut, "Ciblage par Raison Sociale"
  SetField ('RCB_VALIDATION', '-');         // par défaut, pas validé

   if GetParamSocSecur('SO_NUMCIBLAGEAUTO', False) then     // si attribution auto du code ciblage
   begin
    bCiblageAuto        := True;
//    CompteurCiblage     := GetParamSocSecur('SO_COMPTEURCIBLAGE', False);
//    Inc(CompteurCiblage);
//    setField('RCB_CODECIBLAGE', IntToStr(CompteurCiblage));
      SetControlEnabled('RCB_CODECIBLAGE', False);
      SetFocusControl('RCB_LIBELLE');
   end;

end ;

procedure TOM_CIBLAGE.OnDeleteRecord ;
var
   SuppCible : Boolean;

begin
  Inherited ;
  SuppCible := False;
   //si aucun ciblage
  If not ExisteSql('SELECT 1 FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="'+CodeCiblage+'"') then
    SuppCible           := True
  else
  begin
    //si le ciblage à été validé
    if ExisteSql('SELECT 1 FROM CIBLAGE WHERE RCB_CODECIBLAGE="'+CodeCiblage+'" AND RCB_VALIDATION="X"') then
    begin
      LastError         := -1;
      {3, 'Suppression impossible : ciblage validé'}
      LastErrorMsg := TraduireMemoire(TexteMessage[3]);
    end
    else
      {1  'Confirmation'}
      {2, 'Cette cible contient des éléments déjà sélectionnés, voulez-vous vraiment la supprimer ?'}
      if HShowMessage ('0;' + TraduireMemoire(TexteMessage[1]) + ';' + TraduireMemoire(TexteMessage[2])+ ';Q;YN;N;N;','','') = MrYes then
         SuppCible      := True
      else
         LastError      := -1;
  end;

  if SuppCible then
  begin
     ExecuteSql('DELETE FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="'+CodeCiblage+'"');
     exit;
  end;


end ;

procedure TOM_CIBLAGE.OnUpdateRecord ;
var
   BoolMode             : boolean;
   New_Mode             : String;
   NumCompteur          : String;

begin
  Inherited ;

  if bCiblageAuto then
  begin
    if (GetParamSocSecur('SO_PREFIXEAUTOCIBLAGE', False, True)) and (GetField('RCB_CODECIBLAGE') = '') then
      CodeCiblage       := AttribNewCode('CIBLAGE', 'RCB_CODECIBLAGE', 17, V_PGI.User, GetParamSocSecur('SO_COMPTEURCIBLAGE', '', True), '')
    else
      CodeCiblage       := AttribNewCode('CIBLAGE', 'RCB_CODECIBLAGE', 0, '', GetParamSocSecur('SO_COMPTEURCIBLAGE', '', True), '');

    SetField('RCB_CODECIBLAGE', CodeCiblage);
    NumCompteur         := ExtraitChronoCode(CodeCiblage);
    SetParamSoc('SO_COMPTEURCIBLAGE', NumCompteur);
    
  end;


  //si le ciblage à été valider
  if ExisteSql('SELECT 1 FROM CIBLAGE WHERE RCB_CODECIBLAGE="'+CodeCiblage+'" AND RCB_VALIDATION="X"') then
  begin
    LastError           := -1;
    LastErrorMsg        := 'Ciblage validé ! Modification impossible';
  end;

  BoolMode              := False;
  New_Mode              := GetField('RCB_MODECIBLAGE');
  //si le mode a changé
  if Anc_Mode <> New_Mode then
  begin
    If (Anc_Mode = '001') And (New_Mode = '002') then BoolMode := True;
    If (Anc_Mode = '002') And (New_Mode = '001') then BoolMode := True;
    If (Anc_Mode = '003') then BoolMode := true;
    if Anc_Mode = '' then BoolMode := False;
  end;

  {1  'Confirmation'}
  {4, 'Des adresses risquent d''être enlevées du ciblage'}
  If (BoolMode) and (HShowMessage ('0;' + TraduireMemoire(TexteMessage[1]) + ';' + TraduireMemoire(TexteMessage[4]) + ';Q;YN;N;N;','','') = MrYes) then
  begin
    if New_Mode = '002' then
      ExecuteSQL ('DELETE FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="'+CodeCiblage+'" AND RVB_TIERS<>"XXX"')
    else
       ExecuteSQL ('DELETE FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="'+CodeCiblage+'" AND RVB_SUSPECT<>"XXX"');
    if (New_Mode = '002') And (GetField('RCB_TYPECIBLAGE') = '002') then
       ExecuteSQL ('DELETE FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="'+CodeCiblage+'" AND RVB_CONTACT<>"XXX"');
  end;

//  if (Anc_Mode = '' )and (GetParamSocSecur('SO_NUMCIBLAGEAUTO', False)) then     // si attribution auto du code ciblage et nouveau ciblage
//    SetParamSoc('SO_COMPTEURCIBLAGE', GetField('RCB_CODECIBLAGE'));

  if BAvert then
    PGIBox(MsgAvertEnseigne, TFFiche(Ecran).Caption);

  if (Getfield('RCB_RESPONSABLE') <>'') then
      begin
      if not( ExisteRessource (Getfield('RCB_RESPONSABLE'))) then
      begin
        SetFocusControl('RCB_RESPONSABLE');
        Lasterror:=5;
        LastErrorMsg:=TexteMessage[LastError];
        exit ;
      end;
      end;

end ;


procedure TOM_CIBLAGE.OnAfterUpdateRecord ;
begin
  Inherited ;
  end ;

procedure TOM_CIBLAGE.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_CIBLAGE.OnLoadRecord ;
var
  StrSql                : string;
begin
  Inherited ;
   CodeCiblage          := GetField('RCB_CODECIBLAGE');
   Anc_Mode             := '';

   If ExisteSQL('SELECT 1 FROM CIBLAGE WHERE RCB_CODECIBLAGE="'+CodeCiblage+'"') then
   begin
      SetControlEnabled('RCB_TYPECIBLAGE',False);
      Anc_Mode          := GetField('RCB_MODECIBLAGE');

      if (Anc_Mode = '002') and (Assigned(Rb2)) then
        Rb2.Checked     := True
      else if (Anc_Mode = '003') and (Assigned(Rb3)) then
        Rb3.Checked     := True;
   end;

   StrSql               := 'SELECT 1 FROM CIBLAGE WHERE RCB_CODECIBLAGE="'+CodeCiblage+'" AND RCB_VALIDATION="X"';
   if ExisteSQL(StrSql) then
   begin
      SetControlVisible('BDELETE', False);
   end;

end ;

procedure TOM_CIBLAGE.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;


procedure TOM_CIBLAGE.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_CIBLAGE.OnCancelRecord ;
begin
  Inherited ;
end ;






procedure TOM_CIBLAGE.BtUserOnClick(Sender: TObject);
begin
  SetField('RCB_RESPONSABLE', VH_RT.RTResponsable);
end;



(*procedure TOM_CIBLAGE.ThResponsableOnChange(Sender: TObject);
var
   TobRessource         : Tob;

begin
  TobRessource          := Tob.Create('RESSOURCE', nil, -1);
  TobRessource.InitValeurs;
  TobRessource.SelectDB('"'+GetControlText('RCB_RESPONSABLE')+'"', nil, False);
  //SetControlText('NOMCOMPLETUSER', TobRessource.getvalue('ARS_LIBELLE')+' '+TobRessource.getvalue('ARS_LIBELLE2'));
  // PL le 13/04/07 : pour un affichage plus esthet j'ai mis un TLabel...
  SetControlCaption('NOMCOMPLETUSER', TobRessource.getvalue('ARS_LIBELLE')+' '+TobRessource.getvalue('ARS_LIBELLE2'));

  TobRessource.Free;

end; *)



procedure TOM_CIBLAGE.MajCodeCiblage(Sender: TObject);
begin
  if (GetField('RCB_CODECIBLAGE') = '') or (ds.state <> dsInsert) then exit;
  SetField('RCB_CODECIBLAGE',UpperCase_(GetField('RCB_CODECIBLAGE')));
  if ExisteSQL('SELECT 1 FROM CIBLAGE WHERE RCB_CODECIBLAGE="'+GetField('RCB_CODECIBLAGE')+'"') then
  begin
    PGIInfo('Ce code existe déjà !#10#13 Veuillez essayer un autre code.', 'Avertissement');
    SetField('RCB_CODECIBLAGE', '');
    SetFocusControl('RCB_CODECIBLAGE');
  end;
end;

procedure TOM_CIBLAGE.RbOnClick(Sender: Tobject);
var
  StrRb                 : String;

begin
  With Sender as TRadioButton do
  begin
    StrRb               := Name;
    if StrRb = 'RB1' then
      SetField('RCB_MODECIBLAGE', '001')     //TJA 30/07/2008 Correction SetControlText -> SetField
    else if StrRb = 'RB2' then
      SetField('RCB_MODECIBLAGE', '002')
    else
      SetField('RCB_MODECIBLAGE', '003');
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : TJA
Créé le ...... : 03/07/2008
Modifié le ... :   /  /    
Description .. : sur elipsis, ouverture du mul de recherche de ressource 
Suite ........ : pour plus d'informations
Suite ........ : sinon F5 pour recherche sur début code
Suite ........ : CTRL-F5 pour recherche sur le libellé
Mots clefs ... : 
*****************************************************************}
procedure TOM_CIBLAGE.ThResponsableOnElipsisclick(Sender: Tobject);

begin
  SetControlText('RCB_RESPONSABLE',DispatchRecherche(Nil, 3, 'ARS_RESSOURCE=' + GetCOntrolText('RCB_RESPONSABLE'), '', ''));
end;

Initialization
  registerclasses ( [ TOM_CIBLAGE ] ) ;
end.
