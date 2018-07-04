{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 23/06/2006
Modifié le ... : 08/06/2007
Description .. : Source TOF de la FICHE : RTTOPTIERS ()
Suite ........ : Sélection  des éléments pour ciblage sur Tiers
Suite ........ : CIBLAGE TIERS - RAISON SOCIALE
Mots clefs ... : TOF;RTTOPTIERS
*****************************************************************}
Unit UTofRTTOPTIERS_TOF;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     Fe_Main,
{$else}
     eMul,
     MainEagl,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     KPMGUtil,
     ed_tools,
     UtilRT,                    //RTXXWhereConfident, ...
     EntRT,                     //La variable VH_RT
     utilGc,
     UtilMulTrt,                //TraiteEnregMulListe
     utofcomm,                  //Class, MakeWhere
     wcommuns,
     utilselection,             //MulCreerPagesCL
{$IFDEF AFFAIRE}
     UtofAftraducChampLibre,    //Class
{$ENDIF AFFAIRE}
     UTOF ;

Type
{$IFDEF AFFAIRE}
  TOF_RTTOPTIERS = Class (TOF_AFTRADUCCHAMPLIBRE)
{$ELSE}
  TOF_RTTOPTIERS = Class (TTOFCOMM)
{$ENDIF AFFAIRE}
  private
    CodeCiblage : String;
    ModeCiblage : string;

  public

    BAjoute             : TToolBarButton97;
    BEnleve             : TToolBarButton97;
    LaListe             : THGrid;
    BZoom               : TToolbarButton97;
    SqlCible            : String;
    SqlNonCible         : String;
    ConfWhere           : String;


    procedure BAjouteOnClick ( Sender : TObject );
    procedure BEnleveOnClick ( Sender : TObject );
    procedure LaListeOnDblClick ( Sender : TObject );
    procedure BZoomOnClick ( Sender : TObject);

    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (Argument : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Function RTLanceFiche_RTTOPTIERS(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Implementation

{$IFDEF VER150}
  uses
    Variants
    ;
{$ENDIF VER150}

Function RTLanceFiche_RTTOPTIERS(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_RTTOPTIERS.OnArgument (Argument : String ) ;
var
  critere              : string;
  F                    : TFMul;

begin
  Inherited ;
  F                     := TFMul(ecran);
  fTableName            := 'RTCIBLETIERS2';
  fLequel               := 'T_TIERS';
  fPrefixeTable         := 'T';

  SetControlVisible('BOUVRIR', False);

  BAjoute               := TToolBarButton97(getcontrol('BAJOUTE'));
  BAjoute.OnClick       := BAjouteOnClick;

  BEnleve               := TToolBarButton97(GetControl('BENLEVE'));
  BEnleve.OnClick       := BEnleveOnClick;

  LaListe               := THGrid(GetControl('FLISTE'));
  LaListe.OnDblClick    := LaListeOnDblClick;

  BZoom                 := TToolbarButton97(GetControl('BZOOM'));
  If Assigned(BZoom) then
  begin
    BZoom.Visible       := True;
    Bzoom.OnClick       := BZoomOnClick;
  end;

  //mise à jour des champs libres
{$IFNDEF AFFAIRE}
  RTMajChampsLibres(F);
{$ENDIF AFFAIRE}

  MulCreerPagesCL(F,'NOMFIC=GCTIERS');

  Critere               := UpperCase(Trim(Argument));
  CodeCiblage           := READTOKENST(Critere);
  ModeCiblage           := READTOKENST(Critere);

  F.caption             := 'Sélection des ciblages : '+codeciblage;
  UpdateCaption(ecran);

  SetControlVisible('BINSERT', False);
  SetControlVisible('BSELECTALL', True);

  SqlCible              := 'VCM_CODECIBLAGE="' + CodeCiblage + '" ';
  SqlNonCible           := '(T_TIERS NOT IN (SELECT RVB_TIERS FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="' + CodeCiblage + '" ) ' +
                           'AND VCM_CODECIBLAGE="")';

  ConfWhere             := RTXXWhereConfident('CON', True);;
  //FQ10467 gestion des commerciaux
  if not GereCommercial then
  begin
    SetControlVisible ('T_REPRESENTANT', False);
    SetControlVisible ('TT_REPRESENTANT', False);
  end;

end ;

procedure TOF_RTTOPTIERS.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_RTTOPTIERS.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_RTTOPTIERS.OnUpdate ;
begin
  Inherited ;
end ;


procedure TOF_RTTOPTIERS.OnLoad ;
var       // CRM_MNG_012FQ10807_070408
  StrWhere              : String;
  F                     : TFMul;

begin
  Inherited ;
  F                     := TFMul(ecran);
  ReConstruireListe(F);

  if GetCheckBoxState('COCHEE') = CbGrayed then
     StrWhere           := SqlCible + ' OR ' + SqlNonCible
  else if GetCheckBoxState('COCHEE') = CbChecked then
     StrWhere           := SqlCible
  else
     StrWhere           := SqlNonCible;

{CRM_MNG_012FQ10807_070408 }
  SetControlText('XX_WHEREMULTI',MulWhereMultiChoix (F,'RPR',iif( Assigned(TRadioButton(GetControl('MULTIET'))) AND
    (TRadioButton(GetControl('MULTIET')).checked), 'AND','OR')));
{fin CRM_MNG_012FQ10807_070408 }

  SetControlText('XX_WHERE',StrWhere);
  SetControlText('XX_WHERESPE', ConfWhere);


  {$IFDEF EAGLCLIENT}
  AvertirCacheServer('LISTE');
  {$ENDIF}

  if GetControl('NBELEMENTS')<>nil then
    SetControlText('NBELEMENTS', IntToStr(TrouveNbElement(CodeCiblage)));

end ;



procedure TOF_RTTOPTIERS.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_RTTOPTIERS.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_RTTOPTIERS.OnCancel () ;
begin
  Inherited ;
end ;



{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 23/06/2006
Modifié le ... :   /  /
Description .. : ajout d'éléments dans le ciblage
Mots clefs ... :
*****************************************************************}
procedure TOF_RTTOPTIERS.BAjouteOnClick(Sender: TObject);
var
  F                     : TFmul;
  TobElement            : Tob;
  i                     : integer;

begin

  F := Tfmul (ecran);
  if(F.FListe.NbSelected=0) and (not F.FListe.AllSelected) then
  begin
    PGIInfo('Aucun élément sélectionné');
    exit;
  end;

  TobElement            := Tob.Create('Les elements', nil, -1);
  TraiteEnregMulListe(F, 'T_TIERS', 'RTCIBLETIERS2', TobElement, True);
  BEGINTRANS;
  Try
    InitMoveProgressForm(nil, 'Intégration au ciblage ' + CodeCiblage, '', TobElement.Detail.Count, False, True);
    for i := 0 to TobElement.Detail.Count -1 do
    begin
      MoveCurProgressForm(TobElement.Detail[i].GetValue('T_TIERS') + ' -- ' + TobElement.Detail[i].GetValue('T_LIBELLE'));
      CreationVisa (CodeCiblage, TobElement.Detail[i].GetValue('T_TIERS'), TobElement.Detail[i].GetValue('T_NATUREAUXI'), '001', '0');
    end;
    COMMITTRANS;
    FiniMoveProgressForm;
    TobElement.Free;
  except
    ROLLBACK;
    FiniMoveProgressForm;
  end;
  F.BChercheClick(Ecran);
  F.FListe.ClearSelected;

end;



{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 23/06/2006
Modifié le ... :   /  /
Description .. : Enlever un élément de la liste de ciblage
Mots clefs ... :
*****************************************************************}
procedure TOF_RTTOPTIERS.BEnleveOnClick(Sender: TObject);
var
  F                     : TFmul;
  TobElement            : Tob;
  i                     : integer;

begin
  F := Tfmul (ecran);
  if(F.FListe.NbSelected=0) and (not F.FListe.AllSelected) then
  begin
    PGIInfo('Aucun élément sélectionné');
    exit;
  end;

  BEGINTRANS;
  Try
    TobElement        := Tob.Create('Les elements', nil, -1);
    InitMoveProgressForm(nil, 'Suppression d''élément(s) du ciblage ' + CodeCiblage, '', TobElement.Detail.Count, True, True);
    TraiteEnregMulListe(F, 'T_TIERS', 'RTCIBLETIERS2', TobElement, True);
    for i := 0 to TobElement.Detail.Count -1 do
    begin
      MoveCurProgressForm(TobElement.Detail[i].GetValue('T_TIERS') + ' -- ' + TobElement.Detail[i].GetValue('T_LIBELLE'));
      ExecuteSql('DELETE FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="' + CodeCiblage +
                 '" AND RVB_TIERS="' + TobElement.Detail[i].GetValue('T_TIERS')+ '"');
    end;
    TobElement.Free;
    COMMITTRANS;
    FiniMoveProgressForm;
  except
    ROLLBACK;
    FiniMoveProgressForm;
  end;
  F.BChercheClick(Ecran);
  F.FListe.ClearSelected;

end;

procedure TOF_RTTOPTIERS.LaListeOnDblClick(Sender: TObject);
var
  StrTiers              : String;
begin
  if VarIsNull(GetField('T_TIERS')) then
     exit;
  BEGINTRANS;
  try
      if getField('VCM_CODECIBLAGE') <> '' then
       ExecuteSql('DELETE FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="'+CodeCiblage+
                  '" AND RVB_TIERS="'+GetField('T_TIERS')+'"')
    else
       CreationVisa (CodeCiblage, GetField('T_TIERS'), GetField('T_NATUREAUXI'), '001', '0');
    COMMITTRANS;
    StrTiers              := Getfield('T_TIERS');
    TFMul(ecran).BChercheClick(Ecran);
    GotoNewLine(StrTiers);
  except
    ROLLBACK;
  end;
end;


procedure TOF_RTTOPTIERS.BZoomOnClick(Sender: TObject);
begin
  if VarIsNull(GetField('T_TIERS')) then
    exit;

  if GetField('T_NATUREAUXI') = 'CLI' then     // Si fiche Suspect
    V_PGI.DispatchTT(8,taConsult, GetField('T_TIERS'), '', 'MONOFICHE;ACTION=CONSULTATION;T_NATUREAUXI=CLI')
  else if GetField('T_NATUREAUXI')='PRO' then    // si fiche prospect
    V_PGI.DispatchTT(8,taConsult, GetField('T_TIERS'), '', 'MONOFICHE;ACTION=CONSULTATION;T_NATUREAUXI=PRO');


end;

Initialization
  registerclasses ( [ TOF_RTTOPTIERS ] ) ;
end.
