{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 04/07/2012
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTLANCEANALYSE ()
Mots clefs ... : TOF;BTLANCEANALYSE
*****************************************************************}
Unit BTLANCEANALYSE_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul,
{$ENDIF}
     uTob,
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1,
     HMsgBox,
     HTB97,
     UTOF,
     UTOFBTANALDEV,
     SaisUtil;


Type
  TOF_BTLANCEANALYSE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

  private
    Btvalider         : TToolbarButton97;
    RBDocument        : TRadioButton;
    RBCourant         : TRadioButton;
    GRP_Analyse       : TGroupBox;
    GRP_Courant       : TGroupBox;
    LLibelleArt       : THLabel;
    FTitrePiece       : TGroupBox;
    //
    DEV               : RDEVISE;
    //
    TobParam          : TOB;
  	TOBTiers          : TOB;
    TobAffaire        : TOB;
    TobPiece          : TOB;
    TobArticle        : TOB;
    TobOuvrage        : TOB;
    TOBPieceTrait     : TOB;
    TOBPorcs          : TOB;
    TOBBases          : TOB;
    TOBBasesL         : TOB;
    //
    ChkCotraitant     : THCheckBox;
    //
    FlipFlop          : boolean;
    //
    TAnalyse          : TsrcAnal;
    //
    TypeCotrait       : string;
    Parametre         : string;
    //
    procedure OnValideClick(Sender : Tobject);
    procedure OnCourantClick(sender: Tobject);
    procedure OnDocumentClick(sender: Tobject);
    procedure OnCotraitantClick(Sender: TObject);
    procedure GetObjects;
    procedure SetScreenEvents;
    procedure VisibleScreenObject;
    Procedure ChargementEntete;
    procedure ValideAnalyse;
    function ValideRentabilite : boolean;
    //
    function RecupInfoParagraphe: String;
    //

  end ;

Implementation
Uses CalcOLEGenericBTP,
     FactureBtp,
     FactCalc,
     FactUtil,
     SimulRentabDoc,
     uCotraitance,
     EntGC;

procedure TOF_BTLANCEANALYSE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTLANCEANALYSE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTLANCEANALYSE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTLANCEANALYSE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTLANCEANALYSE.OnArgument (S : String ) ;
begin
  Inherited ;

  FlipFlop := false;

  if LaTOB <> nil then
  begin
    TobParam      := LaTob;
    TobAffaire    := TOB(LaTob.data);
    TobTiers      := TOB(TobAffaire.data);
    TobPiece      := TOB(TobTiers.data);
    TobOuvrage    := TOB(TobPiece.data);
    TobArticle    := TOB(TobOuvrage.Data);
    TobPorcs      := TOB(TobArticle.Data);
    if TobPorcs   <> nil  then TOBPieceTrait := Tob(TOBPorcs.data);
  end;
  //
  DEV.Code := TOBPIECE.GetValue('GP_DEVISE');
  GetInfosDevise(DEV);
  DEV.Taux := TobParam.GetDouble('TAUXDEV');
  //
  GetObjects;
  //
  SetScreenEvents;
  //
  VisibleScreenObject;
  //
  ChargementEntete;

end ;

Procedure TOF_BTLANCEANALYSE.GetObjects;
begin

  if assigned(Getcontrol('Bvalider'))         then BtValider      := TToolbarButton97(ecran.FindComponent('BValider'));
  if assigned(Getcontrol('RBDocument'))       then RBDocument     := TRadioButton(ecran.FindComponent('RBDocument'));
  if assigned(Getcontrol('RBCourant'))        then RBCourant      := TRadioButton(ecran.FindComponent('RBCourant'));
  if assigned(Getcontrol('TBPP_PIECEG'))      then FTitrePiece    := TGroupBox(Ecran.FindComponent('TBPP_PIECEG'));
  if Assigned(GetControl('CHKCOTRAITANT'))    then ChkCotraitant  := THCheckBox(Ecran.FindComponent('CHKCOTRAITANT'));
  If Assigned(GetControl('GRP_ANALYSE'))      then GRP_Analyse    := TGroupBox(Ecran.FindComponent('GRP_ANALYSE'));
  If Assigned(GetControl('GRP_COURANT'))      then GRP_Courant    := TGroupBox(Ecran.FindComponent('GRP_COURANT'));
  if Assigned(GetControl('LLIBELLEART'))      then LLibelleArt    := THLabel(Ecran.FindComponent('LLIBELLEART'));

end;

procedure TOF_BTLANCEANALYSE.VisibleScreenObject;
begin


  //if assigned(Btvalider) then Btvalider.Visible := False;

  SetControlVisible('BTDelete',False);
  SetControlVisible('BTInsert',false);
  SetControlVisible('BDefaire',False);
  SetControlVisible('BImprimer',False);
  SetControlVisible('HelpBtn',False);

end;

procedure TOF_BTLANCEANALYSE.SetScreenEvents;
begin

  if assigned(Btvalider)  then BtValider.OnClick := OnValideClick;
  if assigned(RBDocument) then RBDocument.OnClick := OnDocumentClick;
  if assigned(RBCourant)  then RBCourant.OnClick := OnCourantClick;
  if assigned(ChkCotraitant) then ChkCotraitant.OnClick := OnCotraitantClick;

end;

procedure TOF_BTLANCEANALYSE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTLANCEANALYSE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTLANCEANALYSE.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTLANCEANALYSE.OnValideClick(sender : Tobject);
begin
  Inherited ;

  if TobParam.GetString('CODETYPE') = 'A' then
  begin
    ValideAnalyse
  end Else
  begin
    TobParam.SetBoolean('OKVALIDE',ValideRentabilite);
  end;
end;

Procedure TOF_BTLANCEANALYSE.ValideAnalyse;
Var TheTitre    : String;
begin

  TheTitre := TOBPiece.GetValue('GP_NUMERO');

  if TOBPiece.getValue('GP_AFFAIRE') <> '' then
	  TheTitre := TheTitre + ' Affaire : '+ BTPCodeAffaireAffiche(TOBPIece.GetValue('GP_AFFAIRE')) + ' ' + TOBaffaire.Getvalue('AFF_LIBELLE')
  else
  	TheTitre := TheTitre + ' Tiers : '+ TOBPiece.getValue('GP_TIERS') + ' ' + TOBTiers.Getvalue('T_LIBELLE');

  if RBDocument.Checked then
  begin
    if TypeCotrait = '' then
      EntreeAnalyseDocument(TAnalyse, TOBArticle, TOBPiece, TOBOUvrage, ['LIBELLE=' + Thetitre], 1)
    else
      EntreeAnalyseDocument(TAnalyse, TOBArticle, TOBPiece, TOBOUvrage, ['LIBELLE=' + Thetitre, TypeCotrait], 2);
  end
  else if RBcourant.Checked then
  begin
    Thetitre := LLibelleArt.Caption;
    if TypeCotrait = '' then
      EntreeAnalyseDocument(TAnalyse, TOBArticle, TOBPiece, TobOuvrage, ['LIBELLE=' + TheTitre, Parametre], 2)
    else
      EntreeAnalyseDocument(TAnalyse, TOBArticle, TOBPiece, TobOuvrage, ['LIBELLE=' + TheTitre, TypeCotrait, Parametre], 3);
  end;

end ;

function TOF_BTLANCEANALYSE.ValideRentabilite : boolean;
Var indice : Integer;
begin
  Result := false;
  if Entree_Simulation(TOBArticle, TOBPiece, TOBOUvrage,TOBPorcs,TOBBases,TOBBasesL, 'Suivi de rentabilité du document', taModif, DEV, RBCourant.Checked, ChkCotraitant.Checked) then
  begin
    Result := True;
    ReinitMontantPieceTrait (TOBPiece,TOBaffaire,TOBPieceTrait);
    ZeroFacture (TOBpiece);
    ZeroMontantPorts (TOBPorcs);
    for indice := 0 to TOBPiece.detail.count -1 do ZeroLigneMontant (TOBPiece.detail[indice]);
  end;


end;

Function TOF_BTLANCEANALYSE.RecupInfoParagraphe : String;
Var Indice      : Integer;
    TobLoc      : TOB;
    TypeInfo    : String;
    LibCourant  : String;
Begin

  Indice := TobParam.GetInteger('GSLig');
  TOBLoc := TOBPIece.Detail[Indice];

  LibCourant  := TOBLOC.GetString('GL_LIBELLE');

  if IsDebutParagraphe (TOBLOC) then
  begin
    Result    := Result + 'DEBUT=' + intTostr(Indice);
    TAnalyse  := SrcPar;
    TypeInfo  := ' Paragraphe';
  end
	Else if IsFinParagraphe (TOBLOC) then
  begin
    Result    := Result + 'FIN=' + intTostr(Indice);
    TAnalyse  := SrcPar;
    TypeInfo  := ' Paragraphe';
  end
  Else
  begin
    Result    := Result + 'CURLIG=' + intToStr(Indice);
    TAnalyse  := SrcOuv;
    if (TOBLOC.GetValue('GL_TYPEARTICLE') = 'ART') or (TOBLOC.GetValue('GL_TYPEARTICLE') = 'ARP') then
      TypeInfo := ' Article'
    else if (TOBLOC.GetValue('GL_TYPEARTICLE') = 'OUV') then
      TypeInfo := ' Ouvrage'
    else if (TOBLOC.GetValue('GL_TYPEARTICLE') = 'PRE') then
      TypeInfo := ' Prestation'
    else if (TOBLOC.GetValue('GL_TYPEARTICLE') = 'FRS') then
      TypeInfo := ' Frais';
  end;

  if llibelleArt <> nil then
  begin
    if Typeinfo <> '' then LLibelleArt.Caption := TypeInfo + ' : ' + LibCourant
    else LLibelleArt.Caption := LibCourant;
  end;

end;

procedure TOF_BTLANCEANALYSE.OnCourantClick(sender : Tobject);
begin

  //TAnalyse := SrcPar;
  
  Parametre := RecupInfoParagraphe;

  //Affichage des informations de la ligne courante...
  if GRP_courant <> nil then
  begin
    if TobParam.GetString('CODETYPE') = 'A' then
    begin
      ecran.Height := ecran.Height + grp_courant.Height;
      GRP_Courant.Visible := True;
    end;
  end;

  if BtValider <> nil then Btvalider.Visible := True;

end;

procedure TOF_BTLANCEANALYSE.OnDocumentClick(sender : Tobject);
begin

  Parametre := '';

  if GRP_COURANT <> nil then
  begin
    if TobParam.GetString('CODETYPE') = 'A' then
    begin
      ecran.Height := ecran.Height - grp_courant.Height;
      GRP_Courant.Visible := False;
    end;
  end;

  TAnalyse := SrcDoc;

  if BtValider <> Nil then Btvalider.Visible := True;


end;

procedure Tof_BTLANCEANALYSE.OnCotraitantClick(Sender : TObject);
begin

  TypeCotrait := '';

  if ChkCotraitant.Checked then TypeCotrait := 'COTRAITANT=X';

end;

procedure TOF_BTLANCEANALYSE.ChargementEntete;
Var CodeAffaire : String;
    Titrefiche  : String;
begin

  ecran.Height := ecran.Height - grp_courant.Height;

  Titrefiche := GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'), 'GPP_LIBELLE');
  Titrefiche := Titrefiche + ' N° ' + Tobpiece.getstring('GP_NUMERO');
  TitreFiche := Titrefiche + ' du ' + Tobpiece.getstring('GP_DATEPIECE');
  FTitrePiece.Caption := TitreFiche;

  SetControlText('TGP_TIERS',     Tobpiece.getstring('GP_TIERS')+ ' ' + TobTiers.getstring('T_LIBELLE'));

  if Tobpiece.getstring('GP_AFFAIRE1') <> '' then
    CodeAffaire := Tobpiece.getstring('GP_AFFAIRE1');
  if Tobpiece.getstring('GP_AFFAIRE2') <> '' then
    CodeAffaire := codeAffaire + '-' + Tobpiece.getstring('GP_AFFAIRE2');
  if Tobpiece.getstring('GP_AFFAIRE3') <> '' then
    CodeAffaire := codeAffaire + '-' + Tobpiece.getstring('GP_AFFAIRE3');
  if Tobpiece.getstring('GP_AVENANT') <> '' then
    CodeAffaire := codeAffaire + '-' + Tobpiece.getstring('GP_AVENANT');

  SetControlText('TGP_AFFAIRE',  CodeAffaire + ' - ' + TobAffaire.GetString('AFF_LIBELLE'));

  //SetControlText('TGP_NUMERO',    'N° ' + Tobpiece.getstring('GP_NUMERO'));
  SetControlText('TGP_REFINTERNE', Tobpiece.getstring('GP_REFINTERNE'));

  GRP_Analyse.Caption := TOBParam.getString('TYPE');

  if TobParam.GetString('CODETYPE') = 'A' then
    RBCourant.Caption := 'Elément courant'
  else
    RBCourant.caption := 'Par Paragraphe';

end;

Initialization
  registerclasses ( [ TOF_BTLANCEANALYSE ] ) ;
end.

