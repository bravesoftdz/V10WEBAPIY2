{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 26/06/2006
Modifié le ... : 08/06/2007
Description .. : Source TOF de la FICHE : RTTOPTIERSCON ()
Suite ........ : CIBLAGE TIERS - CONTACT
Mots clefs ... : TOF;RTTOPTIERSCON
*****************************************************************}
Unit utofRTTOPTIERSCON_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,Fe_main,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
     Maineagl,
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
     UtilRT,
     EntRT,                     //La variable VH_RT
     utilGC,
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
  TOF_RTTOPTIERSCON = Class (TOF_AFTRADUCCHAMPLIBRE)
{$ELSE}
  TOF_RTTOPTIERSCON = Class (TTOFCOMM)
{$ENDIF AFFAIRE}
  private
    CodeCiblage : String;
    ModeCiblage : string;

  public
    BAjoute             : TToolBarButton97;
    BEnleve             : TToolBarButton97;
    BZoom               : TToolbarButton97;
    LaListe             : THGrid;
    SqlCible            : String;
    SqlNonCible         : String;
    ConfWhere           : String;

    procedure BAjouteOnClick ( Sender : TObject );
    procedure BEnleveOnClick ( Sender : TObject );
    procedure LaListeOnDblClick ( Sender : TObject );
    procedure BZoomOnClick (Sender : Tobject);

    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (Argument : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Function RTLanceFiche_RTTOPTIERSCON(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Implementation

{$IFDEF VER150}
  uses
    Variants
    ;
{$ENDIF VER150}

Function RTLanceFiche_RTTOPTIERSCON(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_RTTOPTIERSCON.OnArgument (Argument : String ) ;
var
   critere              : string;
   F                    : TFMul;

begin
  Inherited ;
  F                     := TFMul(ecran);
  fTableName            := 'RTCIBLETIERSCON';
  fLequel               := 'C_AUXILIAIRE';
  fPrefixeTable         := 'C';

  SetControlVisible('BOUVRIR', False);

  BAjoute               := TToolBarButton97(getcontrol('BAJOUTE'));
  BAjoute.OnClick       := BAjouteOnClick;

  BEnleve               := TToolBarButton97(GetControl('BENLEVE'));
  BEnleve.OnClick       := BEnleveOnClick;

  LaListe               := THGrid(GetControl('FLISTE'));
  LaListe.OnDblClick    := LaListeOnDblClick;

  BZoom                 := TToolbarButton97(GetControl('BZOOM'));
  if Assigned(BZoom) then
  begin
    BZoom.Visible       := True;
    BZoom.OnClick       := BZoomOnClick;
  end;

  //reprise des paramètres pour le lancement de la fiche
{$IFNDEF AFFAIRE}
  RTMajChampsLibres(F);
  RTMajChampsLibresContact(F);
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
  SqlNonCible           := '(VCM_NUMCONTACT NOT IN (SELECT RVB_CONTACT || ";" || CAST(RVB_NUMEROCONTACT AS VARCHAR(5)) FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="' + CodeCiblage + '") ' +
                           'AND VCM_CODECIBLAGE="")';

  ConfWhere             := RTXXWhereConfident('CON', True);;

  //FQ10467 gestion des commerciaux
  if not GereCommercial then
  begin
    SetControlVisible ('T_REPRESENTANT', False);
    SetControlVisible ('TT_REPRESENTANT', False);
  end;


end ;


procedure TOF_RTTOPTIERSCON.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_RTTOPTIERSCON.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_RTTOPTIERSCON.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_RTTOPTIERSCON.OnLoad ;
var
  F                     : TFMul;
 // CRM_MNG_012FQ10807_070408
  StrWhere,whereMulti   : String;

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

  SetControlText('XX_WHERE', StrWhere);
  SetControlText('XX_WHERESPE', ConfWhere);

{CRM_MNG_012FQ10807_070408 }
  whereMulti:='';
  whereMulti:=MulWhereMultiChoix (F,'RPR',iif( Assigned(TRadioButton(GetControl('MULTIET'))) AND
    (TRadioButton(GetControl('MULTIET')).checked), 'AND','OR'));
{fin CRM_MNG_012FQ10807_070408 }


  SetControlText('XX_WHEREMULTI',whereMulti);
  {$IFDEF EAGLCLIENT}
  AvertirCacheServer('LISTE');
  {$ENDIF}

  if GetControl('NBELEMENTS')<>nil then
    SetControlText('NBELEMENTS', IntToStr(TrouveNbElement(CodeCiblage)));

end ;


procedure TOF_RTTOPTIERSCON.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_RTTOPTIERSCON.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_RTTOPTIERSCON.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_RTTOPTIERSCON.BAjouteOnClick(Sender: TObject);
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
  TraiteEnregMulListe(F, 'C_AUXILIAIRE', 'RTCIBLETIERSCON', TobElement, True);
  BEGINTRANS;
  Try
    InitMoveProgressForm(nil, 'Intégration au ciblage ' + CodeCiblage, '', TobElement.Detail.Count, False, True);
    for i := 0 to TobElement.Detail.Count -1 do
    begin
      MoveCurProgressForm(TobElement.Detail[i].GetValue('C_AUXILIAIRE') + ' -- ' + TobElement.Detail[i].GetValue('T_LIBELLE'));
      CreationVisa (CodeCiblage, TobElement.Detail[i].GetValue('C_AUXILIAIRE'), TobElement.Detail[i].GetValue('C_NATUREAUXI'), '002', TobElement.Detail[i].GetValue('C_NUMEROCONTACT'));
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




procedure TOF_RTTOPTIERSCON.BEnleveOnClick(Sender: TObject);
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
    TraiteEnregMulListe(F, 'C_AUXILIAIRE', 'RTCIBLETIERSCON', TobElement, True);
    for i := 0 to TobElement.Detail.Count -1 do
    begin
      MoveCurProgressForm(TobElement.Detail[i].GetValue('C_AUXILIAIRE') + ' -- ' + TobElement.Detail[i].GetValue('T_LIBELLE'));
      ExecuteSql('DELETE FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="' + CodeCiblage +
                 '" AND RVB_CONTACT="' + TobElement.Detail[i].GetValue('C_AUXILIAIRE')+ '"' +
                 ' AND RVB_NUMEROCONTACT="' + IntToStr(TobElement.Detail[i].GetValue('C_NUMEROCONTACT')) + '"');
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




procedure TOF_RTTOPTIERSCON.LaListeOnDblClick(Sender: TObject);
var
  StrTiers              : String;
begin
  if VarIsNull(GetField('C_AUXILIAIRE')) then
     exit;
  BEGINTRANS;
  try
    if GetField('VCM_CODECIBLAGE') <> '' then
       ExecuteSql('DELETE FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="'+CodeCiblage+
                 '" AND RVB_CONTACT="' + GetField('C_AUXILIAIRE')+ '"' +
                 ' AND RVB_NUMEROCONTACT="' + IntToStr(GetField('C_NUMEROCONTACT')) + '"')
    else
       CreationVisa (CodeCiblage, GetField('C_AUXILIAIRE'), GetField('C_NATUREAUXI'), '002', GetField('C_NUMEROCONTACT'));
    COMMITTRANS;
    StrTiers              := Getfield('C_AUXILIAIRE');
    TFMul(ecran).BChercheClick(Ecran);
    GotoNewLine(StrTiers);
  except
    ROLLBACK;
  end;
end;




procedure TOF_RTTOPTIERSCON.BZoomOnClick(Sender: Tobject);
begin
  if VarIsNull(GetField('C_AUXILIAIRE')) then
    exit;

  if GetField('C_NATUREAUXI') = 'CLI' then     // Si fiche Suspect
//    V_PGI.DispatchTT(8,taConsult, GetField('T_AUXILIAIRE'), '', 'MONOFICHE;ACTION=CONSULTATION;T_NATUREAUXI=CLI')
{CRM_20080722_CD_012;10802_DEB}
//    AGLLanceFiche('GC', 'GCTIERS', '', GetField('C_AUXILIAIRE'), 'MONOFICHE;ACTION=CONSULTATION;T_NATUREAUXI=CLI')
    V_PGI.DispatchTT(28,taConsult, GetField('C_AUXILIAIRE'), '', 'MONOFICHE;T_NATUREAUXI=CLI')
{CRM_20080722_CD_012;10802_FIN}
  else if GetField('C_NATUREAUXI')='PRO' then    // si fiche prospect
{CRM_20080722_CD_012;10802_DEB}
//    AGLLanceFiche('GC', 'GCTIERS', '', GetField('C_AUXILIAIRE'), 'MONOFICHE;ACTION=CONSULTATION;T_NATUREAUXI=PRO')
    V_PGI.DispatchTT(28,taConsult, GetField('C_AUXILIAIRE'), '', 'MONOFICHE;T_NATUREAUXI=PRO');
{CRM_20080722_CD_012;10802_FIN}

end;

Initialization
  registerclasses ( [ TOF_RTTOPTIERSCON ] ) ;
end.
