{***********UNITE*************************************************
Auteur  ...... : GF
Créé le ...... : 24/09/2000
Modifié le ... : 19/02/2001
Description .. : Tom du paramétrage planning graphique & numerique (en
Suite ........ : quantité)
Mots clefs ... : PLANNING
*****************************************************************}

unit UtomBTPARAMPLANNING;

interface
uses
  Classes,
  Controls,
  Forms,
  graphics,
  StdCtrls,
  ComCtrls,
  DbCtrls,
  ParamDBG,
  UTOM,
  utob,
  SysUtils,
  msgUtil,
  //
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  EFiche,
  UtileAGL,
  {$ELSE}
  db,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  Fiche,
  Fe_main,
  {$ENDIF}
  //
  M3FP,
  Paramsoc,
  HCtrls,
  HDB,
  Hqry,
  HTB97,
  AglInit,
  //
  HEnt1;
       
type

  TOM_HRPARAMPLANNING = class(TOM)

  public
    procedure OnArgument(Arguments: string); override;
    procedure OnNewRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnLoadRecord; override;
    procedure Duplication(CleDuplic: string); //REGIS

  private

    //Gestion des variables globales
    CodeParam           : String;
    CodeListe           : String;
    DataType            : String;
    //
    //Gestion des boutons
    BCOULEURLUNDI		  	: TToolbarButton97;
    BCOULEURMARDI			  : TToolbarButton97;
    BCOULEURMERCREDI	  : TToolbarButton97;
    BCOULEURJEUDI			  : TToolbarButton97;
    BCOULEURVENDREDI	  : TToolbarButton97;
    BCOULEURSAMEDI		  : TToolbarButton97;
    BCOULEURDIMANCHE	  : TToolbarButton97;
	  BCOULEURVEILJFERIE  : TToolbarButton97;
	  BCOULEURJFERIE			: TToolbarButton97;
    BDUPLICATION				: TToolbarButton97;
    BFOND						    : TToolbarButton97;
    BPARAMITEM					: TToolbarButton97;
    BPARAMHINT					: TToolbarButton97;
    BPARAMGAPP					: TToolbarButton97;
    BSELECTION					: TToolBarButton97;
    BFONTE    					: TToolbarButton97;

    //Gestions des Libellés
    LFond						  : THEdit;
    LSelection				: THEdit;
    LSamedi						: THEdit;
    LDimanche					: THEdit;
    LLundi						: THEdit;
    LMardi						: THEdit;
    LMercredi					: THEdit;
    LJeudi						: THEdit;
    LVendredi					: THEdit;
    LVeilJFerie				: THEdit;
    LJFerie						: THEdit;

    LFonte						: THLabel;
    THPP_FAMRES       : THLabel;
    THPP_FAMMATGERE   : THLabel;

    //Gestion des combo
    HPP_MODEPLANNING 		: THDBValComboBox;
    HPP_FAMRES					: THDBValComboBox;
    HPP_FAMMATGERE      : THDBMultiValComboBox;
    HPP_CADENCEMENT			: THDBValComboBox;
    HPP_FORMATDATE			: THDBValComboBox;
    HPP_CONTENUITEM			: THDBValComboBox;
  	HPP_CONTENUHINT			: THDBValComboBox;
	  HPP_CONTENUGAPP			: THDBValComboBox;

    LibCONTENUGAPP      : THLabel;

    //Gestion des Zônes de Saisie
    HPP_LIBCOL1			  	: THEdit;
    HPP_LIBCOL2			  	: THEdit;
    HPP_LIBCOL3			  	: THEdit;
    //
    HPP_LIBCOLENT4			: THEdit;
    HPP_LIBCOLENT5			: THEdit;
    HPP_LIBCOLENT6			: THEdit;
    HPP_LIBCOLENT7			: THEdit;
    HPP_LIBCOLENT8			: THEdit;
    HPP_LIBCOLENT9  		: THEdit;

    HPP_NBCOLDIVERS			: THDBSpinEdit;
    HPP_NBLIGDIVERS			: THDBSpinEdit;

    //HPP_PLANNINGTYPEET  : TDBCheckBox;
    HPP_PLANNINGTYPETD  : TDBCheckBox;
    HPP_AFFEVTMAT       : THCheckBox;
    HPP_AFFEVTINT       : THCheckBox;
    HPP_AFFEVTCHA       : THCheckBox;
    HPP_AFFABSSAL       : THCheckBox;
    HPP_AFFACTGRC       : THCheckBox;
    HPP_CONTROLCAL      : THCheckbox;

    HPP_VISUTYPERES     : THCheckBox;

    HEUREDEBUT        : THEdit;
    HEUREFIN	        : THEdit;

    ListeItem					: string;
	  ListeHint					: string;
    ListeGapp					: String;
	  Cadencement 			: String;

    AncValNbCol				: Integer;
    AncValNbLig				: Integer;

    //Procedure propres aux évènements des objets de la fiche
    //Procedure BAPERCU_OnClick(Sender: TObject);
    Procedure BCOULEURDIMANCHE_OnClick(Sender: TObject);
    Procedure BCOULEURJEUDI_OnClick(Sender: TObject);
    Procedure BCOULEURJFERIE_OnClick(Sender: TObject);
    Procedure BCOULEURLUNDI_OnClick(Sender: TObject);
    Procedure BCOULEURMARDI_OnClick(Sender: TObject);
    Procedure BCOULEURMERCREDI_OnClick(Sender: TObject);
    Procedure BCOULEURSAMEDI_OnClick(Sender: TObject);
    Procedure BCOULEURVEILJFERIE_OnClick(Sender: TObject);
    Procedure BCOULEURVENDREDI_OnClick(Sender: TObject);
    Procedure BDUPLICATION_OnClick(Sender: TObject);
    Procedure BFOND_OnClick(Sender: TObject);
    Procedure BFONTE_OnClick(Sender: TObject);
    Procedure BPARAMHINT_OnClick(Sender: TObject);
    Procedure BPARAMITEM_OnClick(Sender: TObject);
    Procedure BPARAMGAPP_OnClick(Sender: TObject);
    Procedure BSELECTION_OnClick(Sender: TObject);

    Procedure HPP_CADENCEMENT_OnChange(Sender: TObject);
    Procedure HPP_CONTENUITEM_OnChange(Sender: TObject);
	  Procedure HPP_CONTENUHINT_OnChange(Sender: TObject);
    Procedure HPP_CONTENUGAPP_OnChange(Sender: TObject);
    Procedure HPP_FAMRES_OnChange(Sender: TObject);
    Procedure HPP_NBCOLDIVERS_OnChange(Sender: TObject);
    Procedure HPP_NBLIGDIVERS_OnChange(Sender: TObject);
    Procedure HPP_PLANNINGTYPETD_OnExit(Sender: TObject);

    procedure GestionParamStandard;
    procedure ControleChamp(Champ, Valeur: String);
    procedure ControleCritere(Valeur: String);
    procedure AffichageParametrePlanningPCA;
    procedure AffichageParametrePlanningPLA;
    procedure AffichageParametrePlanningPMA;
    procedure GetObjects;
    procedure SetScreenEvents;
    function ChargementCadencementParc: String;
    function ChargementCadencement: String;
    procedure ControleCadencement(Cadencement: string);
    procedure InitChampTable;
    function ControleSaisie: Boolean;

  end;

implementation


Procedure TOM_HRPARAMPLANNING.OnArgument(Arguments : string);
var Critere : String;
    Valeur  : String;
    Champ   : String;
    X       : integer;
Begin
  Inherited ;

  AncValNbCol := 1;
  AncValNbLig := 1;

  CodeParam := '';
  CodeListe := '';

  GetObjects;

  //traitement Arguments
	Critere:=(Trim(ReadTokenSt(Arguments)));

  while (Critere <> '') do
    begin
      if Critere <> '' then
      begin
        X := pos (':', Critere) ;
        if x = 0 then
          X := pos ('=', Critere) ;
        if x <> 0 then
        begin
          Champ := copy (Critere, 1, X - 1) ;
          Valeur := Copy (Critere, X + 1, length (Critere) - X) ;
        	ControleChamp(champ, valeur);
				end
      end;
      ControleCritere(Critere);
      Critere := (Trim(ReadTokenSt(Arguments)));
    end;

  SetScreenEvents;

  HPP_MODEPLANNING.enabled := False;
  HPP_MODEPLANNING.Value   := CodEParam;

  if pos(CodeParam,'PLA;PSF;PTA;PTR') > 0 then
     AffichageParametrePlanningPLA //Planning Intervention
  else if pos(CodeParam,'PRA;PAR;PFO') > 0 then
     AffichageParametrePlanningPCA //Planning Chantier
  else if pos(CodeParam,'PMA;PFM;PPA') > 0 then
    AffichageParametrePlanningPMA; //Planning Parc/Matériel

  HPP_LIBCOL1.DataType    := DataType;
  HPP_LIBCOL2.DataType    := DataType;
  HPP_LIBCOL3.DataType    := DataType;
  //
  HPP_LIBCOLENT4.DataType := DataType;
  HPP_LIBCOLENT5.DataType := DataType;
  HPP_LIBCOLENT6.DataType := DataType;
  HPP_LIBCOLENT7.DataType := DataType;
  HPP_LIBCOLENT8.DataType := DataType;
  HPP_LIBCOLENT9.DataType := DataType;

  AvertirTable('BTFAMRES');

  HPP_CONTENUGAPP.Visible := False;
  LibCONTENUGAPP.Visible  := false;
  BPARAMGAPP.Visible      := false;

end;       

procedure TOM_HRPARAMPLANNING.OnNewRecord;
Var Req : String;
begin

  AncValNbCol := 1;
  AncValNbLig := 1;

  InitChampTable;

  GestionParamStandard;

  //Controle si un planning par defaut existe déjà pour ce type de planning...
  Req := 'Select HPP_PLANNINGDEFAUT from HRPARAMPLANNING where HPP_PLANNINGDEFAUT="X"';
  Req := Req + 'AND HPP_MODEPLANNING="' + CodeParam + '"';
  if ExisteSQL(req) then
     begin
     SetField('HPP_PLANNINGDEFAUT', '-');
     SetControlEnabled('HPP_PLANNINGDEFAUT', False);
     end
  else
     SetControlEnabled('HPP_PLANNINGDEFAUT', True);
  //
  BPARAMITEM.Visible := false;
  BPARAMHINT.Visible := false;
  BPARAMGAPP.Visible := false;
  //
  //Chargement des combos de mise en forme
  HPP_CONTENUITEM.ItemIndex := 0;
  HPP_CONTENUHINT.ItemIndex := 0;
  HPP_CONTENUGAPP.ItemIndex := 0;
  //
  //on ne va checker que l'option correspondant au planning que l'on est en train de créer.....
  if pos(CodeParam,'PLA;PSF;PTA;PTR') > 0 then //Planning Intervention
  Begin
    HPP_AFFEVTINT.Checked       := False;
    HPP_AFFEVTMAT.Checked       := False;
    HPP_AFFEVTCHA.Checked       := False;
    HPP_AFFACTGRC.Checked       := True;
    HPP_AFFABSSAL.Checked       := True;
  end
  else if pos(CodeParam,'PRA;PAR;PFO') > 0 then //Planning Chantier
  begin
    HPP_AFFEVTINT.Checked       := True;
    HPP_AFFEVTMAT.Checked       := False;
    HPP_AFFEVTCHA.Checked       := False;
    HPP_AFFACTGRC.Checked       := True;
    HPP_AFFABSSAL.Checked       := True;
  end
  else if pos(CodeParam,'PMA;PFM;PPA') > 0 then //Planning Parc/Matériel
  begin
    HPP_AFFEVTINT.Checked       := False;
    HPP_AFFEVTMAT.Checked       := False;
    HPP_AFFEVTCHA.Checked       := False;
    HPP_AFFACTGRC.Checked       := False;
    HPP_AFFABSSAL.Checked       := False;
  end;

  HPP_CONTROLCAL.Checked        := True;

end;

procedure TOM_HRPARAMPLANNING.OnLoadRecord;
var fonte			: String;
begin

  //if getField('HPP_AXECONTINGENT1')='' then
  //begin
  //  if not (DS.State in [dsInsert, dsEdit]) then DS.edit; // pour passer DS.state en mode dsEdit
  //	SetField('HPP_AXECONTINGENT1','001');
  //end;

  // Gestion du chargement des listes d'items et de Hint
  ListeItem := RechDom('BTCONTENUITEM',HPP_CONTENUITEM.Value, true);
  If ListeItem <> '' then BPARAMITEM.Visible := true;

  ListeHint := RechDom('BTCONTENUHINT',HPP_CONTENUHINT.Value, true);
  If ListeHint <> '' then BPARAMHINT.Visible := true;

  //ListeGapp	:= RechDom('BTCONTENUGAPP',HPP_CONTENUGAPP.Value, true);
 	//If ListeGapp <> '' then BPARAMGAPP.Visible := true;

  // Chargement des couleurs dans la fiche
  AfficheCouleur(LFOND, THEdit(GetControl('HPP_COULEURFOND')));
  AfficheCouleur(LSELECTION, THEdit(GetControl('HPP_COULSELECTION')));
  AfficheCouleur(LSAMEDI, THEdit(GetControl('HPP_COULEURSAMEDI')));
  AfficheCouleur(LDIMANCHE, THEdit(GetControl('HPP_COULDIMANCHE')));
  AfficheCouleur(LLUNDI, THEdit(GetControl('HPP_COULLUNDI')));
  AfficheCouleur(LMARDI, THEdit(GetControl('HPP_COULMARDI')));
  AfficheCouleur(LMERCREDI, THEdit(GetControl('HPP_COULMERCREDI')));
  AfficheCouleur(LJEUDI, THEdit(GetControl('HPP_COULJEUDI')));
  AfficheCouleur(LVENDREDI, THEdit(GetControl('HPP_COULVENDREDI')));
  AfficheCouleur(LJFERIE, THEdit(GetControl('HPP_COULJFERIE')));
  AfficheCouleur(LVEILJFERIE, THEdit(GetControl('HPP_COULVEILJF')));

  //Chargement de la fonte dans la fiche
  fonte := GetField('HPP_FONTCOLONNE');
  LFonte.font.name := fonte;

  HEUREDEBUT.Text := FormatDateTime('hh:mm', Getfield('HPP_HEUREDEBUT'));
  HEUREFIN.Text   := FormatDateTime('hh:mm', Getfield('HPP_HEUREFIN'));

  HPP_NBLIGDIVERS.Text := GetField('HPP_NBLIGDIVERS');
  HPP_NBCOLDIVERS.Text := GetField('HPP_NBCOLDIVERS');
  //
  HPP_LIBCOL1.Text := GetField('HPP_LIBCOL1');
  HPP_LIBCOL2.Text := GetField('HPP_LIBCOL2');
  HPP_LIBCOL2.Text := GetField('HPP_LIBCOL2');
  //
  HPP_LIBCOLENT4.Text := GetField('HPP_LIBCOLENT4');
  HPP_LIBCOLENT5.Text := GetField('HPP_LIBCOLENT5');
  HPP_LIBCOLENT6.Text := GetField('HPP_LIBCOLENT6');
  HPP_LIBCOLENT7.Text := GetField('HPP_LIBCOLENT7');
  HPP_LIBCOLENT8.Text := GetField('HPP_LIBCOLENT8');
  HPP_LIBCOLENT9.Text := GetField('HPP_LIBCOLENT9');
  //
  SetControlEnabled('BFONTE', true);
  SetControlEnabled('BFOND', true);
  SetControlEnabled('BSELECTION', true);
  SetControlEnabled('HPP_AFFDATEGROUP', true);
  SetControlEnabled('HPP_AFFICHEICONE', true);
  SetControlEnabled('HPP_MOUSE', true);
  SetControlEnabled('HPP_FORMEGRAPHIQUE', true);
  SetControlEnabled('HPP_MOVEHORIZONTAL', True);
  SetControlVisible('THPP_NOMBRELIGDIVERS', True);
  SetControlVisible('HPP_NOMBRELIGDIVERS', true);
  SetControlEnabled('HPP_VISULIGNETO', True);

  AncValNbCol := StrToInt(GetField('HPP_NBCOLDIVERS'));
  AncValNbLig := StrToInt(GetField('HPP_NBLIGDIVERS'));

  if Pos(CodeParam, 'PMA;PFM;PPA') > 0 then
  Begin
    if HPP_FAMRES.Value <> '' then
    begin
     THPP_FAMMATGERE.Visible := False;
     HPP_FAMMATGERE.Visible  := False;
    end
    else
    begin
     THPP_FAMMATGERE.Visible := True;
     HPP_FAMMATGERE.Visible  := True;
    end;
  end;

  GestionParamStandard;

end;

procedure TOM_HRPARAMPLANNING.OnUpdateRecord;
Var LibCol	: String;
    CodeCol : String;
    I				: Integer;
begin
  //
  Lasterror := 0;
  //
  if Pos(CodeParam,'PMA;PFM;PPA') > 0 then setfield('HPP_ABREGE', Getfield('HPP_LIBELLE'));

  if GetField('HPP_COULEURFOND') = ''   then SetField('HPP_COULEURFOND','ClWhite');

  if GetField('HPP_COULSELECTION') = '' then SetField('HPP_COULSELECTION','ClBlue');

  if GetField('HPP_COULEURSAMEDI') = '' then SetField('HPP_COULEURSAMEDI','ClRed');

  if GetField('HPP_COULDIMANCHE') = ''  then  SetField('HPP_COULDIMANCHE','ClRed');

  if not ControleSaisie then Exit;

  Setfield('HPP_FAMRES',      HPP_FAMRES.Value);
  Setfield('HPP_FAMMATGERE',  HPP_FAMMATGERE.Value);

  SetField('HPP_HEUREDEBUT',  HEUREDEBUT.text);
  SetField('HPP_HEUREFIN',    HEUREFIN.text);

  For i:=1 To StrToInt(HPP_NBLIGDIVERS.text) do
  Begin
    Codecol := GetControlText('HPP_LIBCOL' + IntToStr(i));
    if Pos(CodeParam, 'PMA;PFM;PPA') > 0 then
      LibCol := LectLibCol('CC', 'BMA', Codecol, 'CC_LIBRE')
    else if Pos(CodeParam, 'PAR') > 0 then
      LibCol := LectLibCol('CC', 'BAR', Codecol, 'CC_LIBRE')
    else
      LibCol := LectLibCol('CC', 'BLI', Codecol, 'CC_LIBRE');
    Setfield('HPP_LIBCOL' + IntToStr(i), CodeCol);
  end;

  For i:=1 To StrToInt(HPP_NBCOLDIVERS.text) Do
  Begin
    CodeCol := GetControlText('HPP_LIBCOLENT' + IntToStr(i+3));
    if Pos(CodeParam, 'PMA;PFM;PPA') > 0 then
      LibCol := LectLibCol('CC', 'BMA', CodeCol, 'CC_LIBRE')
    Else if CodeParam = 'PAR' then
      LibCol := LectLibCol('CC', 'BAR', Codecol, 'CC_LIBRE')
    Else
      LibCol := LectLibCol('CC', 'BLI', CodeCol, 'CC_LIBRE');
    Setfield('HPP_LIBCOLENT' + IntToStr(i+3), CodeCol);
  end;
  //
  SetField('HPP_AXECONTINGENT1',  HPP_CONTENUGAPP.Value);
  SetField('HPP_CONTENUHINT',     HPP_CONTENUHINT.Value);
  SetField('HPP_CONTENUITEM',     HPP_CONTENUITEM.Value);
  //
end;

Procedure TOM_HRPARAMPLANNING.InitChampTable;
begin

  if CodeParam = '' then
     SetField('HPP_MODEPLANNING', 'PLA')
  else
     SetField('HPP_MODEPLANNING', CodeParam);

  // Initilisation des champs
  SetField('HPP_TAILLECOL1', 60);
  //
  SetField('HPP_HAUTLIGNEENT',  18);
  SetField('HPP_HAUTLIGNEDATA', 18);
  //
  SetField('HPP_TAILLECOLONNE', 100);
  SetField('HPP_INTERVALLEDEB', 0);
  SetField('HPP_INTERVALLEFIN', 7);
  SetField('HPP_DUREETIMER',    5);
  //
  Setfield('HPP_COULEURSAMEDI',ColorToString($000000FE));
  Setfield('HPP_COULDIMANCHE', ColorToString($000000FE));
  Setfield('HPP_COULSELECTION',ColorToString($00FE0000));
  Setfield('HPP_COULEURFOND',  ColorToString($00FFFFFE));
  //
  Setfield('HPP_FORMEGRAPHIQUE', 'PGL');
  SetField('HPP_FONTCOLONNE', 'Ms Sans Serif');
  SetField('HPP_AFFDATEGROUP', 'X');
  SetField('HPP_PLANNINGGRAPH', 'X');
  //
  SetField('HPP_VISUTYPERES', '-');
  //
  //if Pos(CodeParam, 'PMA;PFM;PPA') > 0 then
  //  SetField('HPP_AXECONTINGENT1','002')      // Rajout LS Suite plantage en entrée dans le planning
  //else if Pos(CodeParam, 'PRA;PAR') > 0 then
  //  SetField('HPP_AXECONTINGENT1','001')     // Rajout LS Suite plantage en entrée dans le planning
  //else
  //  SetField('HPP_AXECONTINGENT1','001');    // Rajout LS Suite plantage en entrée dans le planning
  //
  SetControlProperty('THPP_TAILLECOLENTETE1','Enabled', true);
  SetControlProperty('THPP_TAILLECOLENTETE2','Enabled', False);
  SetControlProperty('THPP_TAILLECOLENTETE3','Enabled', False);
  SetControlProperty('THPP_TAILLECOLENTETE4','Enabled', true);
  SetControlProperty('THPP_TAILLECOLENTETE5','Enabled', False);
  SetControlProperty('THPP_TAILLECOLENTETE6','Enabled', False);
  SetControlProperty('THPP_TAILLECOLENTETE7','Enabled', False);
  SetControlProperty('THPP_TAILLECOLENTETE8','Enabled', False);
  SetControlProperty('THPP_TAILLECOLENTETE9','Enabled', False);
  //
	SetControlProperty('HPP_TAILLECOLENT1','Enabled', true);
  SetControlProperty('HPP_TAILLECOLENT2','Enabled', False);
  SetControlProperty('HPP_TAILLECOLENT3','Enabled', False);
	SetControlProperty('HPP_TAILLECOLENT4','Enabled', true);
  SetControlProperty('HPP_TAILLECOLENT5','Enabled', False);
  SetControlProperty('HPP_TAILLECOLENT6','Enabled', False);
	SetControlProperty('HPP_TAILLECOLENT7','Enabled', False);
  SetControlProperty('HPP_TAILLECOLENT8','Enabled', False);
  SetControlProperty('HPP_TAILLECOLENT9','Enabled', False);
  //
	SetControlProperty('HPP_REGCOL1','Visible', true);
  SetControlProperty('HPP_REGCOL2','Visible', False);
  SetControlProperty('HPP_REGCOL3','Visible', False);
  //
  SetField('HPP_TAILLECOLENT1', 0);
  SetField('HPP_TAILLECOLENT2', 0);
  SetField('HPP_TAILLECOLENT3', 0);
  //
  SetField('HPP_TAILLECOLENT4', 60);
  SetField('HPP_TAILLECOLENT5', 0);
  SetField('HPP_TAILLECOLENT6', 0);
  SetField('HPP_TAILLECOLENT7', 0);
  SetField('HPP_TAILLECOLENT8', 0);
  SetField('HPP_TAILLECOLENT9', 0);
  //
  SetField('HPP_HEUREDEBUT', '00:00');
  SetField('HPP_HEUREFIN', '00:00');

end;

Function TOM_HRPARAMPLANNING.ControleSaisie : Boolean;
var i         : Integer;
    CodeCol   : String;
    Libcol    : string;
begin

  Result := False;

  // Vérification de la validité des champs
  if (CodeParam = 'PLA') then
  Begin
    if HPP_FAMRES.text = '' then
    begin
      Lasterror := 1;
      AfficheErreur(ecran.Name,'1','Paramètre du Planning');
      SetActiveTabSheet('PGENERAL');
      HPP_FAMRES.SetFocus;
      exit;
    end;
  end;

  if StrToInt(HPP_NBLIGDIVERS.text) = 0 then
     begin
     Lasterror := 12;
     AfficheErreur(ecran.Name,'12','Paramètre du Planning');
     SetActiveTabSheet('PMISEFORME');
     HPP_NBLIGDIVERS.SetFocus;
     exit;
     end;

  if StrToInt(HPP_NBCOLDIVERS.Text) = 0 then
     begin
     Lasterror := 13;
     AfficheErreur(ecran.Name,'13','Paramètre du Planning');
     SetActiveTabSheet('PMISEFORME');
     HPP_NBCOLDIVERS.SetFocus;
     exit;
     end;

  if GetField('HPP_FORMATDATECOL0') = '' then
  begin
     Lasterror := 4;
     AfficheErreur(ecran.Name,'4','Paramètre du Planning');
     SetActiveTabSheet('PGENERAL');
     SetFocusControl('HPP_FORMATDATECOL0');
     exit;
  end;

  //Controle des libellés de la colonne d'entête
  For i:=1 To StrToInt(HPP_NBLIGDIVERS.text) do
  Begin
    Codecol := GetControlText('HPP_LIBCOL' + IntToStr(i));
    LibCol := LectLibCol('CC', CodeListe, Codecol, 'CC_LIBRE');
    if LibCol = '' then
    begin
      Lasterror := 4+i;
      AfficheErreur(ecran.Name, IntToStr(Lasterror),'Paramètre du Planning');
      SetActiveTabSheet('PMISEFORME');
      SetFocusControl('HPP_LIBCOL' + IntToStr(i));
      exit;
    end
  end;

  For i:=1 To StrToInt(HPP_NBCOLDIVERS.text) Do
  Begin
    CodeCol := GetControlText('HPP_LIBCOLENT' + IntToStr(i+3));
    LibCol := LectLibCol('CC', CodeListe, CodeCol, 'CC_LIBELLE');
    if LibCol = '' then
    begin
      Lasterror := 10+i;
      AfficheErreur(ecran.Name, IntToStr(Lasterror),'Paramètre du Planning');
      SetActiveTabSheet('PMISEFORME');
      SetFocusControl('HPP_LIBCOLENT' + IntToStr(I+3));
      exit;
    end;
  end;

  Result := True;

end;

procedure TOM_HRPARAMPLANNING.GestionParamStandard;
Var StSQL 		: string;
    i         : Integer;
    Valeur    : Integer;
    Codecol   : string;
    LibCol    : string;
Begin

  // Gestion planning par defaut
  //Controle si un planning par defaut existe déjà pour ce type de planning...
  SetControlEnabled('HPP_PLANNINGDEFAUT', True);

  StSQL := 'Select HPP_PLANNINGDEFAUT from HRPARAMPLANNING where HPP_PLANNINGDEFAUT="X"';
  StSQL := StSQL + 'AND HPP_MODEPLANNING="' + CodeParam + '"';
  SetControlEnabled('HPP_PLANNINGDEFAUT', not ExisteSQL(StSQL));
  //
  if GetField('HPP_PLANNINGDEFAUT') = '-' then
     SetControlEnabled('BDELETE', True)
  else
     SetControlEnabled('BDELETE', False);

  //Chargement du Cadencement en fonction de la famille de Planning
  Cadencement := '' ;
  if Pos(CodeParam, 'PMA;PFM;PPA') > 0 then
    Cadencement := ChargementCadencementParc
  else
  begin
    Cadencement := ChargementCadencement;
  end;

  ControleCadencement(Cadencement);

  //Affichage des libellé des colonne de tri et d'affichage...
  Valeur := StrToInt(HPP_NBCOLDIVERS.Text);
  For i := 1 to Valeur do
  begin
    CodeCol := GetControlText('HPP_LIBCOL' + IntToStr(i));
    LibCol  := LectLibCol('CC', CodeListe, CodeCol, 'CC_LIBELLE');
    SetControlCaption('THPP_COL' + IntToStr(i) + 'LIB', LibCol );
  end;

  Valeur := StrToInt(HPP_NBCOLDIVERS.Text);
  For i := 1 to Valeur do
  begin
    CodeCol := GetControlText('HPP_LIBCOLENT' + IntToStr(i+3));
    LibCol  := LectLibCol('CC', CodeListe, CodeCol, 'CC_LIBELLE');
    SetControlCaption('THPP_COL' + IntToStr(i+3) + 'LIB', LibCol );
  end;

end;

procedure TOM_HRPARAMPLANNING.Duplication(CleDuplic: string);
var TobForm				: TOB;
    i						: integer;
    HrParamPlanning	: string;
    CC					: Tcontrol;
    CC_RO				: Boolean;
begin

  TobForm := TOB.Create('HRPARAMPLANNING', nil, -1);

  if TobForm = nil then exit;

  if CleDuplic = '' then
     begin
     HrParamPlanning := GetControlText('HPP_PARAMPLANNING');
     If (DS.State <> dsInsert) then TFFiche(Ecran).Bouge(NbInsert);
     if not TobForm.SelectDB('"' + HrParamPlanning + '"', nil) then HrParamPlanning := '';
     end;

  if HrParamPlanning <> '' then
     begin
     for i := 1 to TobForm.NbChamps do
         begin
         CC := TControl(Ecran.findcomponent(TobForm.GetNomChamp(i)));
         {$IFDEF EAGLCLIENT}
         CC_RO := ((CC is TEdit) and (TEdit(CC).ReadOnly));
         {$ELSE}
         CC_RO := (((CC is TEdit) and (TEdit(CC).ReadOnly)) or ((CC is TDBEdit) and (TDBEdit(CC).ReadOnly)));
         {$ENDIF}
         if (CC <> nil) and (not CC_RO) then SetField(TobForm.GetNomChamp(i), TobForm.GetValeur(i));
         if (TobForm.GetNomChamp(i) = 'HPP_PARAMPLANNING') then SetField(TobForm.GetNomChamp(i), '');
         if (TobForm.GetNomChamp(i) = 'HPP_PLANNINGDEFAUT') then SetField(TobForm.GetNomChamp(i), '-');
         end;
     SetControlEnabled('HPP_PARAMPLANNING', True);
     end;

  TobForm.Free;

  SetActiveTabSheet(TFFiche(Ecran).pages.Pages[0].name);

  OnLoadRecord;

  SetFocusControl('HPP_PARAMPLANNING');

end;


function TOM_HRPARAMPLANNING.ChargementCadencementParc : String;
Var StSQL : String;
    Q     : TQuery;
Begin

  Result := '';

  if HPP_FAMRES.Value = '' then Exit;

  StSQL := 'SELECT BFM_RYTHMEGESTION FROM BTFAMILLEMAT WHERE BFM_CODEFAMILLE="' + HPP_FAMRES.value + '"';
      Q := OpenSQL(StSQL, True);

  if not Q.EOF then Result := Q.FindField('BFM_RYTHMEGESTION').asString;

  Ferme(Q);

end;

function TOM_HRPARAMPLANNING.ChargementCadencement : String;
Var StSQL : String;
    Q     : TQuery;
Begin

  if HPP_FAMRES.value = '' then Exit;

  StSQL := 'SELECT HFR_RYTHMEGESTION FROM HRFAMRES WHERE HFR_FAMRES="' + HPP_FAMRES.value + '"';
      Q := OpenSQL(StSQL, True);

  if not Q.EOF then Cadencement := Q.FindField('HFR_RYTHMEGESTION').asString;

  Ferme(Q);

end;

procedure TOM_HRPARAMPLANNING.ControleCadencement(Cadencement : string);
begin

  if cadencement = '' then Cadencement := '005';

  SetField('HPP_CADENCEMENT',Cadencement);
  THValComboBox(GetControl('HPP_CADENCEMENT')).Value := Cadencement;

  if (GetField('HPP_CADENCEMENT') = '001') Or
     (GetField('HPP_CADENCEMENT') = '002') Or
     (GetField('HPP_CADENCEMENT') = '003') then
  begin
    HEUREDEBUT.Visible  := True;
    HEUREFIN.Visible    := True;
    HEUREDEBUT.Text     := FormatDateTime('hh:mm', GetParamSocSecur('SO_HEUREDEB', '00:00'));
    HEUREFIN.Text       := FormatDateTime('hh:mm', GetParamSocSecur('SO_HEUREFIN', '00:00'));
    SetControlVisible('THPP_HEUREDEBUT', True);
    SetControlVisible('THPP_HEUREFIN', True);
    THLabel(GetControl('THPP_FORMATDATECOL0')).Caption := 'Format d''heure';
    THValComboBox(GetControl('HPP_FORMATDATECOL0')).Plus := ' AND CO_LIBRE LIKE "H%"';
  end
  else
  begin
    HEUREDEBUT.Visible := False;
    HEUREFIN.Visible := False;
    SetControlVisible('THPP_HEUREDEBUT', False);
    SetControlVisible('THPP_HEUREFIN', False);
    HeureDebut.Text := FormatDateTime('hh:mm', StrToTime('00:00'));
    HeureFin.Text := FormatDateTime('hh:mm', StrToTime('00:00'));
    THValComboBox(GetControl('HPP_FORMATDATECOL0')).Plus := ' AND CO_LIBRE LIKE "D%"';
  end;

end;

procedure TOM_HRPARAMPLANNING.HPP_PLANNINGTYPETD_OnExit(Sender: TObject);
begin

  IF HPP_PLANNINGTYPETD.checked Then
     begin
     SetControlVisible('HPP_LIBCOL5',true);
     SetControlVisible('TTHPP_LIBCOL5',true);
     end
  else
     begin
     SetControlVisible('HPP_LIBCOL5',false);
     SetControlVisible('TTHPP_LIBCOL5',false);
     end;

end;

procedure TOM_HRPARAMPLANNING.HPP_FAMRES_OnChange(Sender: TObject);
Var Cadencement : string;
begin

  if Pos(CodeParam, 'PMA;PFM;PPA') > 0 then
  begin
     if HPP_FAMRES.Value <> '' then
     begin
       THPP_FAMMATGERE.Visible := False;
       HPP_FAMMATGERE.Visible  := False;
     end
     else
     begin
       THPP_FAMMATGERE.Visible := True;
       HPP_FAMMATGERE.Visible  := True;
     end;
     Cadencement := '005';
  end
  else
     Cadencement := ChargementCadencement;

  if cadencement = '' then Cadencement := '005';

  ControleCadencement(Cadencement);

end;

procedure TOM_HRPARAMPLANNING.HPP_CONTENUITEM_OnChange(Sender: TObject);
Begin

	ListeItem := RechDom('BTCONTENUITEM',HPP_CONTENUITEM.Value, true);

	If ListeItem <> '' then BPARAMITEM.Visible := true

end;

procedure TOM_HRPARAMPLANNING.HPP_CONTENUHINT_OnChange(Sender: TObject);
begin

	ListeHint := RechDom('BTCONTENUHINT',HPP_CONTENUHINT.Value, true);

  If ListeHint <> '' then BPARAMHINT.Visible := true

end;

procedure TOM_HRPARAMPLANNING.HPP_CONTENUGAPP_OnChange(Sender: TObject);
Begin

  //ListeGapp	:= RechDom('BTCONTENUGAPP',HPP_CONTENUGAPP.Value,False);

	//If ListeGapp <> '' then BPARAMGAPP.Visible := true

end;

procedure TOM_HRPARAMPLANNING.BPARAMITEM_OnClick(Sender: TObject);
begin

  {$IFDEF EAGLCLIENT}
  ParamListe(ListeItem, nil, '');
  {$ELSE}
  ParamListe(ListeItem, nil, nil, '');
  {$ENDIF}

end;

procedure TOM_HRPARAMPLANNING.BPARAMHINT_OnClick(Sender: TObject);
begin

  {$IFDEF EAGLCLIENT}
  ParamListe(ListeHint, nil, '');
  {$ELSE}
  ParamListe(ListeHint, nil, nil, '');
  {$ENDIF}

end;

procedure TOM_HRPARAMPLANNING.BPARAMGAPP_OnClick(Sender: TObject);
begin

  {$IFDEF EAGLCLIENT}
  //ParamListe(ListeGapp, nil, '');
  {$ELSE}
  //ParamListe(ListeGapp, nil, nil, '');
  {$ENDIF}

end;

procedure TOM_HRPARAMPLANNING.BFond_OnClick(Sender: TObject);
begin

  SelColorNew(LFOND,THEdit(GetControl('HPP_COULEURFOND')), TForm(Ecran));

end;

procedure TOM_HRPARAMPLANNING.BSELECTION_OnClick(Sender: TObject);
begin

  SelColorNew(LSELECTION,THEdit(GetControl('HPP_COULSELECTION')), TForm(Ecran));

end;

procedure TOM_HRPARAMPLANNING.BCOULEURDIMANCHE_OnClick(Sender: TObject);
begin

  SelColorNew(LDIMANCHE,THEdit(GetControl('HPP_COULDIMANCHE')), TForm(Ecran));

end;

procedure TOM_HRPARAMPLANNING.BCOULEURJEUDI_OnClick(Sender: TObject);
begin

  SelColorNew(LJEUDI,THEdit(GetControl('HPP_COULJEUDI')), TForm(Ecran));

end;

procedure TOM_HRPARAMPLANNING.BCOULEURLUNDI_OnClick(Sender: TObject);
begin

  SelColorNew(LLUNDI,THEdit(GetControl('HPP_COULLUNDI')), TForm(Ecran));

end;

procedure TOM_HRPARAMPLANNING.BCOULEURMARDI_OnClick(Sender: TObject);
begin

  SelColorNew(LMARDI,THEdit(GetControl('HPP_COULMARDI')), TForm(Ecran));

end;

procedure TOM_HRPARAMPLANNING.BCOULEURMERCREDI_OnClick(Sender: TObject);
begin

  SelColorNew(LMERCREDI,THEdit(GetControl('HPP_COULMERCREDI')), TForm(Ecran));

end;

procedure TOM_HRPARAMPLANNING.BCOULEURSAMEDI_OnClick(Sender: TObject);
begin

  SelColorNew(LSAMEDI,THEdit(GetControl('HPP_COULEURSAMEDI')), TForm(Ecran));

end;

procedure TOM_HRPARAMPLANNING.BCOULEURVENDREDI_OnClick(Sender: TObject);
begin

  SelColorNew(LVENDREDI,THEdit(GetControl('HPP_COULVENDREDI')), TForm(Ecran));

end;

procedure TOM_HRPARAMPLANNING.BCOULEURJFERIE_OnClick(Sender: TObject);
begin

  SelColorNew(LJFERIE,THEdit(GetControl('HPP_COULJFERIE')), TForm(Ecran));

end;

procedure TOM_HRPARAMPLANNING.BCOULEURVEILJFERIE_OnClick(Sender: TObject);
begin

  SelColorNew(LVEILJFERIE,THEdit(GetControl('HPP_COULVEILJF')), TForm(Ecran));

end;

procedure TOM_HRPARAMPLANNING.BFONTE_OnClick(Sender: TObject);
begin

	SelFonteNew(LFONTE, THEdit(GetControl('HPP_FONTCOLONNE')), TForm(Ecran));

end;

procedure TOM_HRPARAMPLANNING.BDUPLICATION_OnClick(Sender: TObject);
Begin

	Duplication('');

End;

procedure TOM_HRPARAMPLANNING.HPP_CADENCEMENT_OnChange(Sender: TObject);
begin

   Cadencement := HPP_CADENCEMENT.Value;

	if (Cadencement = '001') Or
	   (Cadencement = '002') Or
 	   (Cadencement = '003') then
     Begin
     HEUREDEBUT.Visible := True;
     HEUREFIN.Visible := True;
     HEUREDEBUT.Text := FormatDateTime('hh:mm', GetParamSocSecur('SO_HEUREDEB', '00:00'));
     HEUREFIN.Text := FormatDateTime('hh:mm', GetParamSocSecur('SO_HEUREFIN', '00:00'));
     SetControlVisible('THPP_HEUREDEBUT',True);
     SetControlVisible('THPP_HEUREFIN',True);
     THValComboBox(GetControl('HPP_FORMATDATECOL0')).Plus := ' AND CO_LIBRE LIKE "H%"';
     end
   else
     Begin
     HEUREDEBUT.Visible := False;
     HEUREFIN.Visible := False;
     HEUREDEBUT.Text := FormatDateTime('hh:mm', StrToTime('00:00'));
     HEUREFIN.Text := FormatDateTime('hh:mm', StrToTime('00:00'));
     SetControlVisible('THPP_HEUREDEBUT',False);
     SetControlVisible('THPP_HEUREFIN',False);
     THValComboBox(GetControl('HPP_FORMATDATECOL0')).Plus := ' AND CO_LIBRE LIKE "D%"';
	  end;

end;

Procedure TOM_HRPARAMPLANNING.HPP_NBCOLDIVERS_OnChange(Sender: TObject);
Var i				: integer;
    Valeur	: Integer;
Begin

  Valeur := StrToInt(HPP_NBCOLDIVERS.Text);

  if Valeur > AncValNbCol then
    Begin
    For i := (AncValNbCol+1) To Valeur do
    Begin
     SetControlProperty('THPP_TAILLECOLENTETE'+ IntToStr(I+3),'Enabled', true);
     SetControlProperty('HPP_TAILLECOLENT'    + IntToStr(I+3),'Enabled', true);
     SetControlProperty('THPP_LIBCOL'         + IntToStr(I+3),'Enabled', true);
     SetControlProperty('HPP_LIBCOLENT'       + IntToStr(I+3),'Enabled', True);
     SetField('HPP_TAILLECOLENT'              + IntToStr(I+3), 60);
    end;
    end
  else
  Begin
    For i := AncValNbCol DownTo (Valeur+1) Do
    Begin
      SetControlProperty('THPP_TAILLECOLENTETE'+ IntToStr(I+3),'Enabled', False);
      SetControlProperty('HPP_TAILLECOLENT'    + IntToStr(I+3),'Enabled', False);
      SetControlProperty('THPP_LIBCOL'         + IntToStr(I+3),'Enabled', False);
      SetControlProperty('HPP_LIBCOLENT'       + IntToStr(I+3),'Enabled', False);
      SetField('HPP_TAILLECOLENT'              + IntToStr(I+3), 1);
      SetControlChecked('THHP_LIBCOL'          + IntToStr(I+3), False);
      SetControlCaption('THPP_COL'             + IntToStr(I+3) + 'LIB', '...' );
      SetField('HPP_LIBCOLENT'                 + IntToStr(I+3),'');
    End;
  end;

  AncValNbCol := Valeur;

end;

procedure TOM_HRPARAMPLANNING.HPP_NBLIGDIVERS_OnChange(Sender: TObject);
Var i				: integer;
    Valeur	: Integer;
begin

	Valeur := StrToInt(HPP_NBLIGDIVERS.Text);

  if Valeur > AncValNbLig then
  Begin
    For i := (AncValNbLig+1) To Valeur do
    Begin
      SetControlProperty('THPP_TAILLECOLENTETE' + IntToStr(I),'Enabled', true);
      SetControlProperty('HPP_TAILLECOLENT'     + IntToStr(I),'Enabled', true);
      SetControlProperty('HPP_LIBCOL'           + IntToStr(I),'Enabled', True);
      SetControlProperty('HPP_REGCOL'           + IntToStr(I),'Enabled', True);
      SetControlProperty('THPP_COL'             + IntToStr(I) + 'LIB','Visible', True);
    end;
  end
  else
  Begin
    For i := AncValNbLig DownTo (Valeur+1) Do
    Begin
      SetControlProperty('THPP_TAILLECOLENTETE' + IntToStr(I),'Enabled', False);
      SetControlProperty('HPP_TAILLECOLENT'     + IntToStr(I),'Enabled', False);
      SetControlProperty('HPP_LIBCOL'           + IntToStr(I),'Enabled', False);
      SetControlProperty('HPP_REGCOL'           + IntToStr(I),'Enabled', False);
      //Reinitialisation des zones de tri
      SetField('HPP_TAILLECOLENT'               + IntToStr(I), 0);
      SetField('HPP_LIBCOL'                     + IntToStr(I),'');
      SetControlChecked('HHP_REGCOL'            + IntToStr(I), False);
      //SetControlProperty('THPP_COL'             + IntToStr(I) + 'LIB','Visible', False);
      SetControlCaption('THPP_COL'              + IntToStr(i) + 'LIB', '...' );
    End;
  end;

  AncValNbLig := Valeur;

end;

Procedure TOM_HRPARAMPLANNING.ControleChamp(Champ : String;Valeur : String);
Begin

  if champ = 'MODEPLANNING' then CodeParam := Valeur;

  if Champ = 'ACTION' then
	   Begin
     if Valeur = 'CREATION' Then
     else if Valeur = 'MODIFICATION' then
     else
  end;

end;

procedure TOM_HRPARAMPLANNING.AffichageParametrePlanningPLA;
begin

  CodeListe := 'BLI';
  DataType  := 'BTLIBCOLPLANNING';

  HPP_FAMRES.DataType         := 'BTFAMRES';

  SetControlVisible('THPP_ABREGE',True);
  SetControlVisible('HPP_ABREGE', True);

  HPP_FAMMATGERE.Visible      := False;
  THPP_FAMMATGERE.visible     := False;

  HPP_PLANNINGTYPETD.Visible  := True;

  HPP_AFFEVTINT.Visible       := False;

  HPP_AFFEVTMAT.Visible       := True;
  HPP_AFFEVTCHA.Visible       := True;
  HPP_AFFACTGRC.Visible       := True;
  HPP_AFFABSSAL.Visible       := True;
  HPP_CONTROLCAL.Visible      := True;

  HPP_CONTENUITEM.plus := ' CC_LIBRE=""';
  HPP_CONTENUHINT.plus := ' CC_LIBRE=""';

  if (CodeParam ='PTR') then
  begin
    HPP_FAMRES.Visible      := false;
    THPP_FAMRES.Visible     := False;
    HPP_VISUTYPERES.Visible := False;
  end
  else
  begin
    HPP_FAMRES.Visible      := True;
    THPP_FAMRES.Visible     := True;
    HPP_VISUTYPERES.Visible := True;
  end;

end;

procedure TOM_HRPARAMPLANNING.AffichageParametrePlanningPMA;
begin

    CodeListe := 'BMA';
    DataType  := 'BTPMACOLPLANNING';

    if CodeParam = 'PFM' then
    begin
      HPP_FAMMATGERE.Plus     := 'BFM_NONGEREPLANNING="X"';
      HPP_FAMMATGERE.Visible  := True;
      THPP_FAMMATGERE.visible := True;
      THPP_FAMMATGERE.Left    := THPP_FAMRES.Left;
      HPP_FAMMATGERE.Left     := HPP_FAMRES.Left;
      THPP_FAMRES.Visible     := False;
      HPP_FAMRES.Visible      := False;
    end
    else
    begin
      HPP_FAMRES.DataType     := 'BTFAMILLEMAT';
      HPP_FAMRES.Plus         := 'BFM_NONGEREPLANNING="X"';
      HPP_FAMRES.Visible      := True;
      THPP_FAMRES.Caption     := 'Famille Matériel :';
      THPP_FAMRES.Visible     := True;
      HPP_FAMMATGERE.Plus     := 'BFM_NONGEREPLANNING="X"';
      HPP_FAMMATGERE.Visible  := True;
      THPP_FAMMATGERE.visible := True;
    end;

    HPP_AFFEVTMAT.Visible       := False;
    HPP_AFFACTGRC.Visible       := False;
    HPP_AFFABSSAL.Visible       := False;
    HPP_VISUTYPERES.Visible     := False;
    HPP_CONTROLCAL.Visible      := True;

    HPP_PLANNINGTYPETD.visible  := False;

    HPP_AFFEVTINT.Visible       := True;
    HPP_AFFEVTCHA.Visible       := True;

    SetControlVisible('THPP_ABREGE',False);
    SetControlVisible('HPP_ABREGE', False);

    THValComboBox(GetControl('HPP_CADENCEMENT')).Plus := ' AND CO_LIBRE LIKE "PMA%"';

    HPP_CONTENUITEM.plus := ' CC_LIBRE="PMA"';
    HPP_CONTENUHINT.plus := ' CC_LIBRE="PMA"';

end;

procedure TOM_HRPARAMPLANNING.AffichageParametrePlanningPCA;
begin

  HPP_FAMMATGERE.Visible      := False;
  THPP_FAMMATGERE.visible     := False;

  HPP_PLANNINGTYPETD.Visible  := True;

  HPP_AFFEVTCHA.Visible       := False;
  HPP_AFFEVTMAT.Visible       := True;
  HPP_AFFEVTINT.Visible       := True;

  THValComboBox(GetControl('HPP_CADENCEMENT')).Plus := ' AND CO_LIBRE LIKE "%PCA%"';

  HPP_CONTENUITEM.plus        := ' CC_LIBRE="PCA"';
  HPP_CONTENUHINT.plus        := ' CC_LIBRE="PCA"';

  HPP_CONTROLCAL.Visible      := True;

  //Controle des libellés de la colonne d'entête
  if CodeParam = 'PAR' then
  begin
    CodeListe := 'BAR';
    DataType := 'BTPARCOLPLANNING';
    HPP_PLANNINGTYPETD.Visible:= False;
    THPP_FAMRES.Visible       := False;
    HPP_FAMRES.Visible        := False;
    HPP_VISUTYPERES.Visible   := False;
    HPP_AFFABSSAL.Visible     := False;
    HPP_AFFACTGRC.Visible     := False;
  end
  else
  begin
    CodeListe := 'BLI';
    DataType := 'BTLIBCOLPLANNING';
    HPP_PLANNINGTYPETD.Visible:= True;
    HPP_VISUTYPERES.Visible   := True;
    HPP_AFFABSSAL.Visible     := True;
    HPP_AFFACTGRC.Visible     := True;
    THPP_FAMRES.Visible       := True;
    HPP_FAMRES.Visible        := True;
    if CodeParam = 'PFO' then
    Begin
      THPP_FAMRES.Caption     := 'Fonction :';
      HPP_FAMRES.DataType     := 'AFLFONCTION';
      HPP_FAMRES.PLUS         := ' AND (SELECT BFO_GEREPLANNING FROM BFONCTION WHERE BFO_FONCTION=AFO_FONCTION)="X"';

      HPP_VISUTYPERES.Visible := False;
    end
    else
    begin
      THPP_FAMRES.Caption     := 'Type Ressource :';
      HPP_FAMRES.DataType     := 'BTFAMRES';
      HPP_VISUTYPERES.Visible := True;
    end;
  end;

end;


Procedure TOM_HRPARAMPLANNING.ControleCritere(Valeur : String);
Begin

end;

procedure TOM_HRPARAMPLANNING.GetObjects;
begin

  HPP_PLANNINGTYPETD  := TDBCheckBox(Ecran.FindComponent('HPP_PLANNINGTYPETD'));
  //
  HPP_AFFEVTMAT       := THCheckBox(Ecran.FindComponent('HPP_AFFEVTMAT'));
  HPP_AFFEVTINT       := THCheckBox(Ecran.FindComponent('HPP_AFFEVTINT'));
  HPP_AFFEVTCHA       := THCheckBox(Ecran.FindComponent('HPP_AFFEVTCHA'));
  HPP_AFFABSSAL       := THCheckBox(Ecran.FindComponent('HPP_AFFABSSAL'));
  HPP_AFFACTGRC       := THCheckBox(Ecran.FindComponent('HPP_AFFACTGRC'));
  HPP_CONTROLCAL      := THCheckBox(Ecran.FindComponent('HPP_CONTROLCAL'));
  //
  HPP_VISUTYPERES     := THCheckBox(Ecran.FindComponent('HPP_VISUTYPERES'));

  BCouleurLundi       := TToolbarButton97(ecran.FindComponent('BCOULEURLUNDI'));
  BCouleurMardi       := TToolbarButton97(ecran.FindComponent('BCOULEURMARDI'));
  BCouleurMercredi    := TToolbarButton97(ecran.FindComponent('BCOULEURMERCREDI'));
  BCouleurJeudi       := TToolbarButton97(ecran.FindComponent('BCOULEURJEUDI'));
  BCouleurVendredi    := TToolbarButton97(ecran.FindComponent('BCOULEURVENDREDI'));
  BCouleurSamedi      := TToolbarButton97(ecran.FindComponent('BCOULEURSAMEDI'));
  BCouleurDimanche    := TToolbarButton97(ecran.FindComponent('BCOULEURDIMANCHE'));
  BCouleurVeilJFerie  := TToolbarButton97(ecran.FindComponent('BCOULEURVEILJFERIE'));
  BCouleurJFerie      := TToolbarButton97(ecran.FindComponent('BCOULEURJFERIE'));
  BDuplication        := TToolbarButton97(ecran.FindComponent('B_DUPLICATION'));
  BFond               := TToolbarButton97(ecran.FindComponent('BFOND'));

  BFonte      := TToolbarButton97(ecran.FindComponent('BFONTE'));

  BSelection  := TToolbarButton97(ecran.FindComponent('BSELECTION'));

  BPARAMITEM  := TToolbarButton97(ecran.FindComponent('BPARAMITEM'));

  BPARAMHINT  := TToolbarButton97(ecran.FindComponent('BPARAMHINT'));

  BPARAMGAPP  := TToolbarButton97(ecran.FindComponent('BPARAMGAPP'));

  LFond 			:= THEdit(ecran.FindComponent('LFOND'));
	LSelection 	:= THEdit(ecran.FindComponent('LSELECTION'));
  LSamedi 		:= THEdit(ecran.FindComponent('LSAMEDI'));
  LDimanche 	:= THEdit(ecran.FindComponent('LDIMANCHE'));
  LLundi 			:= THEdit(ecran.FindComponent('LLUNDI'));
  LMardi 			:= THEdit(ecran.FindComponent('LMARDI'));
  LMercredi 	:= THEdit(ecran.FindComponent('LMERCREDI'));
  LJeudi 			:= THEdit(ecran.FindComponent('LJEUDI'));
  LVendredi 	:= THEdit(ecran.FindComponent('LVENDREDI'));
  LVeilJFerie := THEdit(ecran.FindComponent('LVEILJFERIE'));
  LJFerie 		:= THEdit(ecran.FindComponent('LJFERIE'));

	LFonte 			      := THLabel(ecran.FindComponent('LFONTE'));
  THPP_FAMRES       := THLabel(ecran.FindComponent('THPP_FAMRES'));
  THPP_FAMMATGERE   := THLabel(ecran.FindComponent('THPP_FAMMATGERE'));

  HPP_MODEPLANNING  := THDBValComboBox(Ecran.FindComponent('HPP_MODEPLANNING'));

  HPP_CADENCEMENT	  := THDBValComboBox(Ecran.FindComponent('HPP_CADENCEMENT'));
  HPP_FORMATDATE	  := THDBValComboBox(Ecran.FindComponent('HPP_FORMATDATECOL0'));

  HPP_CONTENUITEM   := THDBValComboBox(Ecran.FindComponent('HPP_CONTENUITEM'));

  HPP_CONTENUHINT   := THDBValComboBox(Ecran.FindComponent('HPP_CONTENUHINT'));

  HPP_CONTENUGAPP   := THDBValComboBox(Ecran.FindComponent('HPP_AXECONTINGENT1'));
  LibCONTENUGAPP    := THLabel(Ecran.FindComponent('THPC_AXECONTINGENT'));

  HPP_NBCOLDIVERS   := THDBSpinEdit(Ecran.FindComponent('HPP_NBCOLDIVERS'));

  HPP_NBLIGDIVERS   := THDBSpinEdit(Ecran.FindComponent('HPP_NBLIGDIVERS'));

  HPP_LIBCOL1 := THEdit(GetControl('HPP_LIBCOL1'));
  HPP_LIBCOL2 := THEdit(GetControl('HPP_LIBCOL2'));
  HPP_LIBCOL3 := THEdit(GetControl('HPP_LIBCOL3'));
  //
  HPP_LIBCOLENT4 := THEdit(Ecran.FindComponent('HPP_LIBCOLENT4'));
  HPP_LIBCOLENT5 := THEdit(Ecran.FindComponent('HPP_LIBCOLENT5'));
  HPP_LIBCOLENT6 := THEdit(Ecran.FindComponent('HPP_LIBCOLENT6'));
  HPP_LIBCOLENT7 := THEdit(Ecran.FindComponent('HPP_LIBCOLENT7'));
  HPP_LIBCOLENT8 := THEdit(Ecran.FindComponent('HPP_LIBCOLENT8'));
  HPP_LIBCOLENT9 := THEdit(Ecran.FindComponent('HPP_LIBCOLENT9'));

  HEUREDEBUT  := THEdit(Ecran.FindComponent('HEUREDEBUT'));

  HPP_FAMRES      := THDBValComboBox(Ecran.FindComponent('HPP_FAMRES'));
  HPP_FAMMATGERE  := THDBMultiValComboBox(Ecran.FindComponent('HPP_FAMMATGERE'));

  HEUREFIN 	  := THEdit(Ecran.FindComponent('HEUREFIN'));

end;

procedure TOM_HRPARAMPLANNING.SetScreenEvents;
begin

  BCouleurLundi.onclick       := BCouleurLundi_OnClick;
  BCouleurMardi.onclick       := BCouleurMardi_OnClick;
  BCouleurMercredi.onclick    := BCouleurMercredi_OnClick;
  BCouleurJeudi.onclick       := BCouleurJeudi_OnClick;
  BCouleurVendredi.onclick    := BCouleurVendredi_OnClick;
  BCouleurSamedi.onclick      := BCouleurSamedi_OnClick;
  BCouleurDimanche.onclick    := BCouleurDimanche_OnClick;
  BCouleurVeilJFerie.OnClick  := BCouleurVeilJFerie_OnClick;
  BCouleurJFerie.OnClick      := BCouleurJFerie_OnClick;
  BFonte.onclick              := BFonte_OnClick;
  BFond.onclick               := BFond_OnClick;
  BSelection.onclick          := BSelection_OnClick;
  BPARAMITEM.onclick          := BPARAMITEM_OnClick;
  BPARAMHINT.onclick          := BPARAMHINT_OnClick;
  BDuplication.onclick        := BDuplication_OnClick;
  BPARAMGAPP.onclick          := BPARAMGAPP_OnClick;

  HPP_PLANNINGTYPETD.OnExit   := HPP_PLANNINGTYPETD_OnExit;
  HPP_CADENCEMENT.OnChange    := HPP_CADENCEMENT_OnChange;
  HPP_NBLIGDIVERS.OnChange    := HPP_NBLIGDIVERS_OnChange;
  HPP_NBCOLDIVERS.OnChange    := HPP_NBCOLDIVERS_OnChange;
  HPP_CONTENUGAPP.OnChange    := HPP_CONTENUGAPP_OnChange;
  HPP_CONTENUHINT.OnChange    := HPP_CONTENUHINT_OnChange;
  HPP_CONTENUITEM.OnChange    := HPP_CONTENUITEM_OnChange;

  HPP_FAMRES.OnChange         := HPP_FAMRES_OnChange;

end;

initialization
  RegisterClasses([TOM_HRPARAMPLANNING]);
  //InitPlanning;
end.
