{***********UNITE*************************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/04/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPGLANA ()
Suite ........ :
Suite ........ : JP 01/07/05 : Gestion des caractères Joker : fonctions de
Suite ........ : base définies dans TofMeth
Mots clefs ... : TOF;CPGLANA
*****************************************************************}
Unit uTofCPGLAna ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Windows,
{$IFDEF EAGLCLIENT}
     eMul,
     MainEagl,      // AGLLanceFiche
     eQRS1,         // TFQRS1
{$ELSE}
     DB,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Mul,
     FE_main,       // AGLLanceFiche
     QRS1,          // TFQRS1
{$ENDIF}
     Spin,          // TSpinEdit
     TofMeth,       //
     AGLInit,       // TheData
     CritEdt,       // ClassCritEdt
     Htb97,
     Forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     HZoomSp,
     uTob,         // TOB
     utilPGI,      // TSQLAnaCroise
     Ent1,         // CorrespToCodes
     ParamSoc;     //FP   GetParamSoc

Type
  TOF_CPGLANA = Class (TOF_METH)

    Pages         : TPageControl;
    RGDEVISE      : THRadioGroup;

    Y_SECTION        : THEdit;
    Y_SECTION_       : THEdit;
    Y_GENERAL        : THEdit;
    Y_GENERAL_       : THEdit;
    YDATECOMPTABLE   : THEdit;
    YDATECOMPTABLE_  : THEdit;
    Y_NUMEROPIECE    : THEdit;
    Y_NUMEROPIECE_   : THEdit;
    Y_REFINTERNE     : THEdit;
    CPTEXCEPT        : THEdit;
    SECTIONEXCEPT    : THEdit;

    YQUALIFPIECE     : THMultiValComboBox;

    G_NATUREGENE     : THValComboBox;

    Y_AXE            : THValComboBox;
    MODESELECT       : THValComboBox;
    Y_DEVISE         : THValComboBox;
    Y_ETABLISSEMENT  : THValComboBox;
    Y_EXERCICE       : THValComboBox;

    // Compléménts
    YVALIDEES        : TCheckBox;


    // Correspondances
    CORRESP          : THValComboBox;
    CORRESPDE        : THValComboBox;
    CORRESPA         : THValComboBox;

    // Tables Libres
    TABLELIBRE       : THValComboBox;
    LIBREDE          : THEdit;
    LIBREA           : THEdit;

    // Onglet Rupture
    TABRUPTURES      : TTabSheet;
    RUPTURE          : THRadioGroup; // Sans ou Avec
    RUPTURETYPE      : THRadioGroup; // Groupes, Tables Libres, Correspondance
    RUPGROUPES       : TGroupBox;    // GroupBox visible en fonction de RuptureType
    RUPLIBRES        : TGroupBox;    // GroupBox visible en fonction de RuptureType
    RUPCORRESP       : TGroupBox;    // GroupBox visible en fonction de RuptureType
    NIVORUPTURE      : TSpinEdit;    // Niveau de la rupture 1,2,3
    CPTCorrespExist  : TCheckBox;
    CPTLibresExist   : TCheckBox;

    // Composants Paramètres Développeur
    LIBYQUALIFPIECE  : THEdit;
    AFFDEVISE        : THEdit;
    DEVPIVOT         : THEdit;
    YEXERCICE        : THEdit;
    DATEDEBUTEXO     : THEdit;
    CUMULAU          : THEdit;

    GLTYPE           : THEdit; // 0, 1, 2
    GLXPARY          : THEdit; // - , X
    XX_RUPTURE       : THEdit; // Nom du champ de la rupture paramètrable
    AVECRUPTURE      : THEdit; // AVEC, SANS
    AVECRUPTYPE      : THEdit; // RUPGROUPES, RUPLIBRES, RUPCORRESP
    AVECNIVORUPTURE  : THEdit; // 1, 2, 3
    JOINTUREOPTION    : THEdit; // Gestion des ecritures sans extourne
    YSANSEXTOURNE : TCheckBox;

    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure ChargementCritEdt        ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure MySelectFiltre           ; override ;

    procedure OnClickRupture       ( Sender : TObject );
    procedure OnClickRuptureType   ( Sender : TObject );

    procedure OnChangeG_NATUREGENE ( Sender : TObject );

    procedure OnChangeY_Axe        ( Sender : TObject );
    procedure OnChangeY_Exercice   ( Sender : TObject );
    procedure OnChangeCorresp      ( Sender : TObject );
    procedure OnChangeTableLibre   ( Sender : TObject );
    procedure OnEnterQualifPiece   ( Sender : TObject ); // SBO 12/10/2004
    procedure FTimerTimer          ( Sender : TObject );
    procedure onY_SECTIONKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  private
    FInGLType : Integer;
    FStNumAxe : String;

    fOnSaveKeyDownSection : procedure(Sender: TObject; var Key: Word; Shift: TShiftState) of object;
    fOnSaveKeyDownSection_ : procedure(Sender: TObject; var Key: Word; Shift: TShiftState) of object;
    function  RecupCriteresRuptures : string;
    procedure RemplitPages;
    procedure RemplitParametresDEV;
    procedure GenereCumulGL;
  end ;

function CPLanceFiche_CPGLANA        : string ;
function CPLanceFiche_CPGLGENEPARANA : string ;
function CPLanceFiche_CPGLANAPARGENE : string ;

Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  CPProcGen,
  {$ENDIF MODENT1}
  uLibCalcEdtCompta,  // FTobCumulAu
  uMultiDossierUtil,  // RequeteMultiDossier
  uLibExercice,       // CInitComboExercice
  uLibWindows;        // TraductionTHMultiValComboBox

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/04/2004
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
function CPLanceFiche_CPGLANA : string ;
begin
  Result := AGLLanceFiche('CP', 'CPGLANA', '', '', '');
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/04/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function CPLanceFiche_CPGLGENEPARANA : string ;
begin
  Result := AGLLanceFiche('CP', 'CPGLGENEPARANA', '', '', '');
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/04/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function CPLanceFiche_CPGLANAPARGENE : string ;
begin
  Result := AGLLanceFiche('CP', 'CPGLANAPARGENE', '', '', '');
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/04/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLANA.OnArgument (S : String ) ;
begin
  Inherited ;
  // GCO - 11/07/2006 - FQ 18363
  PAvances := TTabSheet(GetControl('AVANCES', true));
  Pavances.TabVisible := True;
  TFQRS1(Ecran).CritAvancesVisibled := True;

  GLTYPE  := THEdit(GetControl('GLTYPE', True));
  GLXPARY := THEdit(GetControl('GLXPARY', True));

  // Détermination du type de l'état
  if Ecran.Name = 'CPGLANA' then
  begin // Grand livre analytique
    FInGLType := 0;
    Ecran.HelpContext := 7421000 ; // GCO - 04/06/2007 - FQ 15520
  end
  else
  begin
    // Grand livre général par analytique ou Grand livre analytique par général
    Y_GENERAL  := THEdit(GetControl('Y_GENERAL', True));
    Y_GENERAL_ := THEdit(GetControl('Y_GENERAL_', True));

    if Ecran.Name = 'CPGLGENEPARANA' then
    begin // Grand livre général par analytique
      FInGLType := 1;
      Ecran.HelpContext := 7436100 ; // GCO - 04/06/2007 - FQ 15520
      G_NATUREGENE := THValComboBox(GetControl('G_NATUREGENE', True));
      G_NATUREGENE.OnChange := OnChangeG_NATUREGENE;
    end
    else
    begin // Grand livre analytique par général
      FInGLType := 2;
      Ecran.HelpContext := 7439100 ; // GCO - 04/06/2007 - FQ 15520
    end;
  end;

  Pages            := TPageControl(GetControl('PAGES', True));
  TABRUPTURES      := TTabSheet(GetControl('TABRUPTURES', True));

  Y_SECTION        := THEdit(GetControl('YSECTION', True));    {FP 12/10/2005 Pour générer la condition SQL dans Delphi}
  Y_SECTION_       := THEdit(GetControl('YSECTION_', True));   {FP 12/10/2005}

  { CA - 05/10/2005 - Pour la saisie des axes structurés }
  if assigned(Y_SECTION) then
  begin
    fOnSaveKeyDownSection := Y_SECTION.OnKeyDown;
    Y_SECTION.OnKeyDown := onY_SECTIONKeyDown;
  end;
  if assigned(Y_SECTION_) then
  begin
    fOnSaveKeyDownSection_ := Y_SECTION_.OnKeyDown;
    Y_SECTION_.OnKeyDown := onY_SECTIONKeyDown;
  end;

  YDATECOMPTABLE   := THEdit(GetControl('YDATECOMPTABLE', True));
  YDATECOMPTABLE_  := THEdit(GetControl('YDATECOMPTABLE_', True));
  Y_NUMEROPIECE    := THEdit(GetControl('Y_NUMEROPIECE', True));
  Y_NUMEROPIECE_   := THEdit(GetControl('Y_NUMEROPIECE_', True));
  Y_REFINTERNE     := THEdit(GetControl('Y_REFINTERNE', True));

  CPTEXCEPT        := THEdit(GetControl('CPTEXCEPT', False));
  SECTIONEXCEPT    := THEdit(GetControl('SECTIONEXCEPT', False));

  YVALIDEES        := TCheckBox(GetControl('YVALIDEES', True));

  RUPTURE          := THRadioGroup(GetControl('RUPTURE', True));
  RUPTURETYPE      := THRadioGroup(GetControl('RUPTURETYPE', True));
  RGDEVISE         := THRadioGroup(GetControl('RGDEVISE', True));

  RUPCORRESP       := TGroupBox(GetControl('RUPCORRESP', True));
  RUPLIBRES        := TGroupBox(GetControl('RUPLIBRES', True));
  RUPGROUPES       := TGroupBox(GetControl('RUPGROUPES', True));
  NIVORUPTURE      := TSpinEdit(GetControl('NIVORUPTURE', True));

  YQUALIFPIECE     := THMultiValComboBox(GetControl('YQUALIFPIECE', True));

  CORRESP          := THValComboBox(GetControl('CORRESP', True));
  CORRESPDE        := THValComboBox(GetControl('CORRESPDE', True));
  CORRESPA         := THValComboBox(GetControl('CORRESPA', True));
  Y_AXE            := THValComboBox(GetControl('YAXE', True));        {FP 12/10/2005}
  MODESELECT       := THValComboBox(GetControl('MODESELECT', True));
  Y_EXERCICE       := THValComboBox(GetControl('Y_EXERCICE', True));
  Y_ETABLISSEMENT  := THValComboBox(GetControl('Y_ETABLISSEMENT', True));
  Y_DEVISE         := THValComboBox(GetControl('Y_DEVISE', True));
  TABLELIBRE       := THValComboBox(Getcontrol('TABLELIBRE', True));
  LIBREDE          := THEdit(GetControl('LIBREDE', True));
  LIBREA	   := THEdit(GetControl('LIBREA', True));
  CPTCorrespExist  := TCheckBox(GetControl('CPTCORRESPEXIST', True));
  CPTLibresExist   := TCheckBox(GetControl('CPTLIBRESEXIST', True));

  // Composants Paramètres développeur
  AFFDEVISE        := THEdit(GetControl('AFFDEVISE', True));
  DEVPIVOT         := THEdit(GetControl('DEVPIVOT', True));
  LIBYQUALIFPIECE  := THEdit(GetControl('LIBYQUALIFPIECE', True));
  CUMULAU          := THEdit(GetControl('CUMULAU', True));
  YEXERCICE        := THEdit(GetControl('YEXERCICE', True));
  DATEDEBUTEXO     := THEdit(GetControl('DATEDEBUTEXO', True));
  XX_RUPTURE       := THEdit(GetControl('XX_RUPTURE', True));
  AVECRUPTYPE      := THEdit(GetControl('AVECRUPTYPE', True));
  AVECRUPTURE      := THEdit(GetControl('AVECRUPTURE', True));
  AVECNIVORUPTURE  := THEdit(GetControl('AVECNIVORUPTURE', True));
  JOINTUREOPTION   := THEdit(GetControl('JOINTUREOPTION', True));
  YSANSEXTOURNE    := TCheckBox(GetControl('YSANSEXTOURNE', True));

  // Gestion de la combo YEXERCICE
  CInitComboExercice(Y_EXERCICE);
  Y_Exercice.OnChange     := OnChangeY_Exercice;
  if ( CtxPCl in V_PGI.PgiContexte ) and  ( VH^.CPExoRef.Code <>'' ) then
    Y_EXERCICE.Value := CExerciceVersRelatif(VH^.CPExoRef.Code)
  else
    Y_EXERCICE.Value := CExerciceVersRelatif(VH^.Entree.Code) ;

  Y_Axe.OnChange         := OnChangeY_Axe;
  Corresp.OnChange       := OnChangeCorresp;
  TableLibre.OnChange    := OnChangeTableLibre;
  Rupture.OnClick        := OnClickRupture;
  RuptureType.OnClick    := OnClickRuptureType;
  FTimer.OnTimer         := FTimerTimer;
  YQUALIFPIECE.OnEnter   := OnEnterQualifPiece ; // SBO 12/10/2004

  Y_Axe.ItemIndex := 0;
  OnChangeY_Axe( nil);

  if FInGLType = 1 then
  begin
    G_NatureGene.ItemIndex := 0;
    OnChangeG_NatureGene( nil );
  end;
  {JP 28/06/06 : FQ 16149 : géré dans TOF_METH
  PositionneEtabUser(Y_ETABLISSEMENT);}
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/04/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLANA.OnNew ;
begin
  inherited ;
  // ----------------- Init des composants si pas de Filtre DEFAUT -------------
  if FFiltres.Text = '' then
  begin
    Y_Axe.ItemIndex := 0;
    ModeSelect.ItemIndex := 1;
    YQualifPiece.Text := 'N;';
    {JP 04/10/05 : FQ 16149 : si l'établissement est positionné (PositionneEtabUser),
                   il est préférable d'éviter de le "dépositionner" }
    if (Y_Etablissement.Enabled) and (Y_Etablissement.ItemIndex < 0) then
      Y_Etablissement.ItemIndex := 0;

    Y_Devise.ItemIndex        := 0;
  end;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/04/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLANA.OnDisplay () ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/09/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPGLANA.ChargementCritEdt;
begin
  inherited;

  if (TheData <> nil) and (TheData is ClassCritEdt) then
  begin
    // GCO - 16/09/2005 - FQ 16644
    if ClassCritEdt(TheData).CritEdt.Gl.Axe <> '' then
      Y_AXE.Value := ClassCritEdt(TheData).CritEdt.Gl.Axe;

    Case FInGLType of
      0 : begin //gl analytique
            Y_SECTION.Text :=  ClassCritEdt(TheData).CritEdt.Cpt1;
            Y_SECTION_.Text :=  ClassCritEdt(TheData).CritEdt.Cpt2;
          end;
      1 : begin //gl générale par analytique
            Y_General.Text  := ClassCritEdt(TheData).CritEdt.Cpt1;
            Y_General_.Text := ClassCritEdt(TheData).CritEdt.Cpt2;
            Y_SECTION.Text :=  ClassCritEdt(TheData).CritEdt.sCpt1;
            Y_SECTION_.Text :=  ClassCritEdt(TheData).CritEdt.sCpt2;
          end;
      2 : begin //gl anaytique par générale
            Y_SECTION.Text :=  ClassCritEdt(TheData).CritEdt.Cpt1;
            Y_SECTION_.Text :=  ClassCritEdt(TheData).CritEdt.Cpt2;
            Y_General.Text  := ClassCritEdt(TheData).CritEdt.sCpt1;
            Y_General_.Text := ClassCritEdt(TheData).CritEdt.sCpt2;
          end;
    end;

    // Exercice
    Y_Exercice.Value := CExerciceVersRelatif(ClassCritEdt(TheData).CritEdt.Exo.Code);
    // Date de Début de l'édition
    YDateComptable.Text  := DateToStr(ClassCritEdt(TheData).CritEdt.Date1);
    // Date de Fin de l'édition
    YDateComptable_.Text := DateToStr(ClassCritEdt(TheData).CritEdt.Date2);

    // GCO - 13/09/2005 - FQ 10637
    if ClassCritEdt(TheData).CritEdt.GL.NumPiece1 > 0 then
      Y_NumeroPiece.Text  := IntToStr(ClassCritEdt(TheData).CritEdt.GL.NumPiece1);

    if ClassCritEdt(TheData).CritEdt.GL.NumPiece2 > 0 then
      Y_NumeroPiece_.Text := IntToStr(ClassCritEdt(TheData).CritEdt.GL.NumPiece2);

    // Référence Interne
    Y_RefInterne.Text := ClassCritEdt(TheData).CritEdt.ReferenceInterne;
    // Qualifpièce
    YQualifPiece.Text := ClassCritEdt(TheData).CritEdt.QualifPiece;
    // Etablissement
    Y_Etablissement.Value := ClassCritEdt(TheData).CritEdt.Etab;
    // Devise
    Y_Devise.Value := ClassCritEdt(TheData).CritEdt.DeviseSelect;
    // Ecritures validées ( '' = Grayed, OUI = Checked sinon not Checked
    if (ClassCritEdt(TheData).CritEdt.Valide = '') then
      YValidees.State := cbGrayed
    else
      YValidees.Checked := (ClassCritEdt(TheData).CritEdt.Valide = 'OUI');

    TheData := nil;
  end;
end;
////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/04/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLANA.OnLoad ;
begin
  Inherited ;
  //
  RemplitPages;
  //
  RemplitParametresDEV;

  // GCO - 12/11/2007 - FQ 21758 Calcul des Cumuls ANO et Cumuls Au du GL
  GenereCumulGL;

end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 03/11/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
function TOF_CPGLANA.RecupCriteresRuptures : string;
begin
  Result := '';

  if Rupture.Value <> 'SANS' then
  begin
    // Rupture sur champ libre généraux
    if (RuptureType.Value = 'RUPLIBRES') and (TableLibre.ItemIndex >= 0) then
    begin
      if CPTLibresExist.Checked then
      begin
        if Trim(LibreDe.Text) <> '' then
          Result := XX_Rupture.Text + '>= "' + LibreDe.Text + '"';

        if Trim(LibreA.Text) <> '' then
          Result := Result + ' AND '+ XX_Rupture.Text + '<= "' + LibreA.Text + '"';
      end;
    end;

    // Rupture sur plan de correspondance
    if (RuptureType.Value = 'RUPCORRESP') and (Corresp.ItemIndex >= 0) then
    begin
      if CPTCorrespExist.Checked then
      begin
        if Trim(CorrespDe.Text) <> '' then
          Result := XX_Rupture.Text + '>= "' + CorrespDe.Text + '"';

        if Trim(CorrespA.Text) <> '' then
          Result := Result + ' AND ' + XX_Rupture.Text + '<= "' + CorrespA.Text + '"';
      end;
    end;
    
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/04/2004
Modifié le ... :   /  /
Description .. : Au lancement de l'état
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLANA.OnUpdate ;
var lStListeExcept     : string;
    lStExcept          : string;
    lStValYQualifPiece : string;
    lStLibYQualifPiece : string;
    lStRecupCriteresRuptures : string;
    lStOrderBy : string;
begin
  inherited ;

  // Sélection du type d'écritures et  Traduction du E_QualifPiece
  TraductionTHMultiValComboBox( YQualifPiece, lStValYQualifPiece, lStLibYQualifPiece, 'Y_QUALIFPIECE', False );
  LibYQualifPiece.Text := lStLibYQualifPiece;

  //YMO (Tâche devt 3963) 07/04/2006 Exclusion des écritures d'extourne
  if YSANSEXTOURNE.Checked then
  begin
      TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL
      + ' AND ('
      + ' (E_QUALIFORIGINE <> "URN" AND E_QUALIFORIGINE <> "EXT")'
      + ' OR (EC_DATEEXT<"'+USDateTime(StrToDate(YDATECOMPTABLE.Text))
      + '" OR EC_DATEEXT>"'+USDateTime(StrToDate(YDATECOMPTABLE_.Text))
      + '") OR Y_TYPEANALYTIQUE = "X") ';

      JOINTUREOPTION.Text:='LEFT JOIN ECRCOMPL ON E_JOURNAL = EC_JOURNAL AND E_EXERCICE = EC_EXERCICE'
      +' AND E_NUMEROPIECE = EC_NUMEROPIECE AND E_DATECOMPTABLE = EC_DATECOMPTABLE'
      +' AND E_QUALIFPIECE = EC_QUALIFPIECE AND E_NUMLIGNE = EC_NUMLIGNE';
  end
  else
      JOINTUREOPTION.Text:='';
  // GCO - 19/05/2006 - FQ 17744
  // Comptes ou sections d'exceptions, séparés par des ',' ou ';'
  if Assigned(CptExcept) then
  begin
    if CptExcept.Text <> '' then
    begin
      lStListeExcept := FindEtReplace(CptExcept.Text, ',', ';', True);
      while (lStListeExcept <> '') do
      begin
        lStExcept := Trim(ReadTokenSt(lStListeExcept));
        if lStExcept <> '' then
          TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + ' AND Y_GENERAL NOT LIKE "' + lStExcept + '%"';
      end;
    end;
  end;  

  if Assigned(SectionExcept) then
  begin
    if SectionExcept.Text <> '' then
    begin
      lStListeExcept := FindEtReplace(SectionExcept.Text, ',', ';', True);
      while (lStListeExcept <> '') do
      begin
        lStExcept := Trim(ReadTokenSt(lStListeExcept));
        if lStExcept <> '' then
          TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + ' AND Y_SECTION NOT LIKE "' + lStExcept + '%"';
      end;
    end;
  end;

  // Récupération des écritures dans la fourchette de date
  TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL +
    ' AND ((Y_DATECOMPTABLE >= "' + UsDateTime(StrToDate(YDATECOMPTABLE.Text)) + '"' +
    ' AND Y_DATECOMPTABLE <= "'+ UsDateTime(StrToDate(YDATECOMPTABLE_.Text)) + '"' +
    ' AND ' + lStValYQualifPiece + ' AND Y_ECRANOUVEAU = "N") OR ' +
  // Récupération systématique de l'A-Nouveau pour provoquer la rupture dans l'état
    '(Y_DATECOMPTABLE = "' + UsDateTime(StrToDate(DATEDEBUTEXO.Text)) + '" AND Y_ECRANOUVEAU = "OAN"))';

  // Condition WHERE des critères de Rupture
  lStRecupCriteresRuptures := RecupCriteresRuptures;
  if lStRecupCriteresRuptures <> '' then
    TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + ' AND ' + lStRecupCriteresRuptures;

  // FP 12/10/2005 Condtion sur l'axe et la fourchette de compte
  TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL +
    ' AND S_AXE="'+Y_Axe.Value+'"'+
    ' AND ' + TSQLAnaCroise.ConditionAxe(Y_Axe.Value) +
    // GCO - 05/06/2007 - FQ 19175
    ' AND ' + ConvertitCaractereJokers(Y_SECTION, Y_SECTION_ , TSQLAnaCroise.ChampSection(Y_Axe.Value));
  // FP 12/10/2005}

  // FP 12/10/2005 Remplit les controles utilisés par le générateur d'état
  SetControlText('CONDITIONAXE', TSQLAnaCroise.ConditionAxe(Y_Axe.Value));
  SetControlText('CHAMPSECTION', TSQLAnaCroise.ChampSection(Y_Axe.Value));
  // FP 12/10/2005

  (*
  // Gestion du V_Pgi.Confidentiel
  if (FInGLType = 0) then
  begin
    FRequeteEtat := FRequeteEtat + ' AND ' + CGenereSQLConfidentiel('S');
  end
  else
  begin
    FRequeteEtat := FRequeteEtat + ' AND ' + CGenereSQLConfidentiel('S') +
                    ' AND ' + CGenereSQLConfidentiel('G');
  end;*)

  // Traitement pour conversion du Relatif vers Exercice
  TFQRS1(Ecran).WhereSQL := CMajRequeteExercice ( Y_EXERCICE.Value, TFQRS1(Ecran).WhereSQL);

  // Gestion de l'ORDER BY
  if (Rupture.ItemIndex = 0) or ((Rupture.ItemIndex <> 0) and (RuptureType.ItemIndex = 0)) then
  begin
    lStOrderBy := ' ORDER BY ';
  end
  else
  begin
    lStOrderBy := ' ORDER BY ' + XX_Rupture.Text + ',' ;
  end;

  case FInGLType of
  {b FP 12/10/2005
  0 : lStOrderBy := lStOrderBy + 'Y_SECTION, Y_AXE, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE, Y_NUMVENTIL';
  1 : lStOrderBy := lStOrderBy + 'Y_GENERAL, Y_SECTION, Y_AXE, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE, Y_NUMVENTIL';
  2 : lStOrderBy := lStOrderBy + 'Y_SECTION, Y_GENERAL, Y_AXE, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE, Y_NUMVENTIL';}
  0 : lStOrderBy := lStOrderBy + TSQLAnaCroise.ChampSection(Y_Axe.Value)+','+
        {' S_AXE,'+ Lek&FP 08/11/05 FQ16997, il faut enlever 1 S_AXE de Select}
        ' Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE, Y_NUMVENTIL';
  1 : lStOrderBy := lStOrderBy + 'Y_GENERAL,'+
        ' '+TSQLAnaCroise.ChampSection(Y_Axe.Value)+','+
        {' S_AXE,'+ Lek&FP 08/11/05 FQ16997, il faut enlever 1 S_AXE de Select}
        ' Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE, Y_NUMVENTIL';
  2 : lStOrderBy := lStOrderBy + TSQLAnaCroise.ChampSection(Y_Axe.Value)+','+
        ' Y_GENERAL,'+
        {' S_AXE,'+ Lek&FP 08/11/05 FQ16997, il faut enlever 1 S_AXE de Select}
        ' Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE, Y_NUMVENTIL';
  {e FP 12/10/2005}
  else
    PgiError('Erreur sur le type du Grand Livre.', 'Attention');
  end;

  TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + lStOrderBy;
end ;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPGLANA.OnDelete ;
begin
  inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/04/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLANA.FTimerTimer(Sender: TObject);
begin
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPGLANA.OnChangeY_Exercice(Sender: TObject);
begin
  CExoRelatifToDates ( Y_EXERCICE.Value, YDATECOMPTABLE, YDATECOMPTABLE_);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/10/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLANA.OnChangeG_NatureGene( Sender : TObject );
begin
  Y_General.Text := '';
  Y_General_.Text := '';

  Y_General.Plus := IIF(G_NatureGene.ItemIndex > 0, ' AND G_NATUREGENE="' + G_NatureGene.Value + '"', '');
  Y_General_.Plus := Y_General.Plus;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... :
Modifié le ... : 04/11/2004
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLANA.OnChangeY_Axe( Sender : TObject );
var lTFB   : TFichierBase ;
begin
  // GCO - 05/06/2006 - FQ 16318
  if Y_Axe.ItemIndex < 0 then
  begin
    FStNumAxe := '';
    Exit ;
  end;
  FStNumAxe := Copy(Y_Axe.Value, 2, 1);

  { FQ 19572 BVE 14.06.07 }
  if not(cLoadfiltre) then
  begin
     Y_Section.Text  := '';
     Y_Section_.Text := '';
  end;
  { END FQ 19572 }

  if FStNumAxe = '1' then
  begin
    Y_Section.DataType  := 'TZSECTION' ;
    Y_Section_.DataType := 'TZSECTION' ;
  end
  else
  begin
    Y_Section.DataType  := 'TZSECTION' + FStNumAxe;
    Y_Section_.DataType := 'TZSECTION' + FStNumAxe;
  end ;

  // 2. Type de généraux
  if FInGLType <> 0 then
  begin
    Y_General.Text  := '';
    Y_General_.Text := '';
    Y_General.DataType  := 'TZGVENTIL' + FStNumAxe ;
    Y_General_.DataType := 'TZGVENTIL' + FStNumAxe ;
  end;

  // 3. Type de Plan de correspondance
  if FInGLType = 1 then
  begin
    if GetParamSocSecur('SO_CORSGE2', False, True) then
      Corresp.plus := 'AND (CO_CODE = "GE1" OR CO_CODE = "GE2")'
    else
      Corresp.plus := 'AND CO_CODE = "GE1"';
  end
  else
  begin
    // Type de plan comptable sur Axe1 au chargement de l'écran
    if GetParamSocSecur('SO_CORSA12', False, True) then
      Corresp.plus := 'AND (CO_CODE = "A11" OR CO_CODE = "A12")'
    else
      Corresp.plus := 'AND CO_CODE = "A11"';
  end;

  // 3. Longueur des sections
  lTFB := AxeToFb('A' + FStNumAxe);
  Y_Section.MaxLength  := VH^.Cpta[lTFB].Lg ;
  Y_Section_.MaxLength := VH^.Cpta[lTFB].Lg ;
  
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLANA.OnChangeCorresp(Sender: TObject);
begin
  if RuptureType.Value <> 'RUPCORRESP' then Exit;
  CorrespToCodes(Corresp, TComboBox(CorrespDe), TComboBox(CorrespA));
  CorrespDe.ItemIndex := 0 ;
  CorrespA.ItemIndex  := CorrespA.Items.Count - 1 ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/04/2004
Modifié le ... : 04/11/2004
Description .. : Changement de la rupture sur les Tables Libres
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLANA.OnChangeTableLibre(Sender: TObject);
begin
  if TableLibre.ItemIndex < 0 then
  begin
    LibreDe.DataType := '';
    LibreA.DataType  := '';
    LibreDe.Text     := '';
    LibreA.Text      := '';
  end
  else
  begin
    if FInGLType = 1 then
    begin
      LibreDe.DataType := 'TZNATGENE' + GetNumTableLibre;
      LibreA.DataType  := 'TZNATGENE' + GetNumTableLibre;
      LibreDe.Text     := '';
      LibreA.Text      := '';
    end
    else
    begin
      LibreDe.DataType := 'TZNATSECT' + GetNumTableLibre;
      LibreA.DataType  := 'TZNATSECT' + GetNumTableLibre;
      LibreDe.Text     := '';
      LibreA.Text      := '';
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/04/2004
Modifié le ... : 04/11/2004
Description .. : Affichage de l'onglet RUPTURES
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLANA.OnClickRupture(Sender: TObject);
begin
  TABRuptures.TabVisible := (Rupture.ItemIndex <> 0);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/04/2004
Modifié le ... : 04/11/2004
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLANA.OnClickRuptureType(Sender: TObject);
begin
  RupGroupes.Visible := (RuptureType.ItemIndex = 0);
  RupLibres.Visible  := (RuptureType.ItemIndex = 1);
  RupCorresp.Visible := (RuptureType.ItemIndex = 2);

  if RuptureType.ItemIndex = 1 then
  begin
    TableLibre.ItemIndex := 0;
    OnChangeTableLibre(nil);
  end;

  if RuptureType.ItemIndex = 2 then
  begin
    if Corresp.Items.Count >= 1 then
      Corresp.ItemIndex := 1
    else
      Corresp.ItemIndex := 0;

    OnChangeCorresp(nil);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/04/2004
Modifié le ... :   /  /
Description .. : Remplissage des zones du TPageControl lors du lancement de
Suite ........ : l'état
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLANA.RemplitPages;
var lStNatureGene : string;
begin

  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if not TestJoker(Y_Section.Text) then begin
    // Remplissage des bornes des sections
    if (Trim(Y_Section.Text) = '') then
      Y_Section.Text := GetMinCompte('SECTION', 'S_SECTION', Y_Axe.Value);

    if (Trim(Y_section_.Text) = '') then
      Y_Section_.Text := GetMaxCompte('SECTION', 'S_SECTION', Y_Axe.Value);
  end;

  // Remplissage des comptes généraux
  if (FInGLType <> 0) then
  begin
    // Nature du Compte Général pour le GL Général par analytique uniquement
    if FInGLType = 1 then
      lStNatureGene := G_NATUREGENE.Value
    else
      lStNatureGene := '';

    {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
    if not TestJoker(Y_GENERAL.Text) then begin
      if Trim(Y_GENERAL.Text) = '' then
        Y_GENERAL.Text := GetMinCompte('GENERAL', 'G_GENERAL', lStNatureGene);

      if Trim(Y_GENERAL_.Text) = '' then
        Y_GENERAL_.Text := GetMaxCompte('GENERAL', 'G_GENERAL', lStNatureGene);
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLANA.RemplitParametresDEV;
var lStChamp : string;
    lExoDate : TExoDate;
begin
  // Affichage en DEVISE ou en EURO
  if (RGDevise.Enabled) and (RGDevise.ItemIndex = 1) then
  begin
    // Affichage en Devise
    AFFDEVISE.Text := 'X';
    DEVPIVOT.Text  := TraduireMemoire('Devise');
  end
  else
  begin
    AFFDEVISE.Text := '-';
    DEVPIVOT.Text  := RechDom('TTDEVISE', V_PGI.DevisePivot, False);
  end;

  // YEXERCICE
  YEXERCICE.Text := CRelatifVersExercice(Y_EXERCICE.Value);

  // Date Début d'exercice
  lExoDate := CtxExercice.QuelExoDate(CRelatifVersExercice(Y_EXERCICE.Value));
  DATEDEBUTEXO.Text := DateToStr(lExoDate.Deb);

  // Date de CUMULAU
  CUMULAU.Text := 'Cumul au ' + DateToStr(StrToDate(YDateComptable.Text) - 1);

  // Enregistrement du Type de Grand Livre dans un critère pour l' état
  GLTYPE.Text := IntToStr( FInGLType );

  // Type d'état X par Y ( Général par analytique OU analytique par Général )
  GLXPARY.Text := IIF( FInGLType = 0, '-', 'X');

  AVECRUPTURE.Text     := 'SANS' ;
  AVECRUPTYPE.Text     := '' ;
  AVECNIVORUPTURE.Text := '0' ;
  XX_RUPTURE.Text      := '' ;

  if Rupture.Value <> 'SANS' then
  begin
    AVECRUPTURE.Text := Rupture.Value;
    AVECRUPTYPE.Text := RuptureType.Value;

    // Rupture sur Numéro de Compte
    if (RuptureType.Value = 'RUPGROUPES') and (NivoRupture.Value > 0) then
    begin
      AVECNIVORUPTURE.Text := IntToStr(NivoRupture.Value);
    end;

    // Rupture sur champ libre généraux
    if (RuptureType.Value = 'RUPLIBRES') and (TableLibre.ItemIndex >= 0) then
    begin
      lStChamp := IIF( FInGLType = 1, 'G_TABLE', 'S_TABLE');
      XX_Rupture.Text := lStChamp + GetNumTableLibre;
    end;

    // Rupture sur plan de correspondance
    if (RuptureType.Value = 'RUPCORRESP') and (Corresp.ItemIndex >= 0) then
    begin
      lStChamp := IIF( FInGLType = 1, 'G_CORRESP', 'S_CORRESP');
      XX_Rupture.Text := lStChamp + IntToStr(Corresp.ItemIndex + 1);
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 12/10/2004
Modifié le ... :   /  /
Description .. : Afin de rendre apparent le focus sur la
Suite ........ : THMultiValComboBax, car en readOnly, on perd le curseur
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLANA.OnEnterQualifPiece(Sender: TObject);
begin
  CSelectionTextControl( Sender ) ;
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPGLANA.OnCancel () ;
begin
  inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPGLANA.OnClose ;
begin
  if Assigned(FTobCumulAnoGL) then FreeAndNil(FTobCumulAnoGL);
  if Assigned(FTobCumulAuGL)  then FreeAndNil(FTobCumulAuGL);
  inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 05/10/2005
Modifié le ... :   /  /
Description .. : Gestion des accès au choix des sections dans les zones de
Suite ........ : saisie des sections
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLANA.onY_SECTIONKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var St : string;
    fb : TFichierBase ;
begin
  St := THCritMaskEdit(Sender).Text;
  fb := AxeToFb(Y_AXE.Value);
  if (Shift = []) and (Key = 187) then
  begin
    Key := 0;
    Y_SECTION_.Text := Y_SECTION.Text;
  end
  else
  if ((Shift=[ssCtrl]) And (Key=VK_F5)) then
  begin
    if (fb in [fbAxe1..fbAxe5]) and
       VH^.Cpta[fb].Structure and
       // GCO - 29/11/2006 - FQ 19175
       ExisteSQL('SELECT SS_AXE FROM STRUCRSE WHERE SS_AXE = "' + FBToAxe(fb) + '"')
    then
    begin
      if ChoisirSousPlan( fb, St , True,taModif) then
      begin
        if ((THCritMaskEdit(Sender) = Y_SECTION_) and EstJoker(St)) then
          Y_SECTION.Text := St
        else
          THCritMaskEdit(Sender).Text := St;
      end;
      Key := 0;
    end;
  end;

  if THCritMaskEdit(Sender).Name = 'YSECTION' then
    fOnSaveKeyDownSection (Sender, Key, Shift)   {FP 12/10/2005}
  else
    fOnSaveKeyDownSection_ (Sender, Key, Shift);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/10/2005
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPGLANA.MySelectFiltre;
var lExoDate : TExoDate;
begin
  inherited;
  lExoDate := CtxExercice.QuelExoDate(CRelatifVersExercice(Y_Exercice.Value));
  if (CtxExercice.QuelExo(YDateComptable.Text, False) <> lExoDate.Code) and
     (CtxExercice.QuelExo(YDateComptable_.Text, False) <> lExoDate.Code) then
  begin
    CExoRelatifToDates(Y_EXERCICE.Value, YDATECOMPTABLE, YDATECOMPTABLE_);
  end;

  // GCO - 30/07/2007 - FQ 21180
  if YQUALIFPIECE.Value = '' then
  begin
    YQualifpiece.SelectAll;
    YQualifpiece.Text := TraduireMemoire('<<Tous>>');
  end;

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/11/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPGLANA.GenereCumulGL;
var lStSelect : HString;
    lStChampSection : HString;
    lStConditionAxe : HString;
    lStWhere        : HString;
    lStTemp         : HString;
    lStGrpBy        : HString;
    lStWhereCommun  : HString;
    lStValYQualifPiece : String;
    lStLibYQualifPiece : String;
    lStRgrp            : String;
    lQuery             : TQuery;
begin
  if FTobCumulAnoGL = nil then
    FTobCumulAnoGL := Tob.Create('', nil, -1)
  else
    FTobCumulAnoGL.ClearDetail;

  if FTobCumulAuGL = nil then
    FTobCumulAuGL := Tob.Create('', nil, -1)
  else
    FTobCumulAuGL.ClearDetail;

  lStChampSection := GetControlText('CHAMPSECTION');
  lStConditionAxe := GetControlText('CONDITIONAXE');

  case FInGLType of

    0 : begin
          lStSelect := 'SELECT ' + lStChampSection + ',';
          lStGrpBy  := 'GROUP BY ' + lStChampSection + ', Y_EXERCICE, Y_ECRANOUVEAU';
        end;

    1 : begin
          lStSelect := 'SELECT Y_GENERAL, ' + lStChampSection + ',';
          lStGrpBy  := 'GROUP BY Y_GENERAL, ' + lStChampSection + ', Y_EXERCICE, Y_ECRANOUVEAU';
        end;

    2 : begin
          lStSelect := 'SELECT ' + lStChampSection + ', Y_GENERAL, ';
          lStGrpBy  := 'GROUP BY ' + lStChampSection + ', Y_GENERAL, Y_EXERCICE, Y_ECRANOUVEAU';
        end;

    else
    begin
      lStSelect := '';
      lStWhere  := '';
      lStGrpBy  := '';
    end;
  end;

  if AffDevise.Text = 'X' then
    lStSelect := lStSelect + ' SUM(Y_DEBITDEV) DEBIT, SUM(Y_CREDITDEV) CREDIT, '
  else
    lStSelect := lStSelect + ' SUM(Y_DEBIT) DEBIT, SUM(Y_CREDIT) CREDIT, ';

  lStSelect := lStSelect + 'Y_EXERCICE, Y_ECRANOUVEAU FROM ANALYTIQ WHERE ' ;





  lStWhere := 'Y_DATECOMPTABLE = "' + UsDateTime(StrToDate(DateDebutExo.Text)) + '" AND ' +
              'Y_ECRANOUVEAU = "OAN" AND Y_QUALIFPIECE = "N"';

    // Where Commun aux 2 requêtes
  lStWhereCommun := '';

  // Gestion de l'axe Choisit
  lStWhereCommun := lStwhereCommun + ' AND '  + lStConditionAxe;

  if Y_Exercice.Value <> '' then
    lStWhereCommun := lStWhereCommun + ' AND Y_EXERCICE = "' + CRelatifVersExercice(Y_EXERCICE.Value) + '"';

  if Y_Etablissement.Value <> '' then
    lStWhereCommun := lStWhereCommun + ' AND Y_ETABLISSEMENT = "' + Y_Etablissement.Value + '"';

  if Y_Devise.Value <> '' then
    lStWhereCommun := lStWhereCommun + ' AND Y_DEVISE = "' + Y_Devise.Value + '"';

  lStTemp := ConvertitCaractereJokers( Y_SECTION, Y_SECTION_, 'Y_SECTION');
  if Trim(lStTemp) <> '' then
    lStWhereCommun := lStWhereCommun + ' AND ' + lStTemp;

  if (FInGLType <> 0) then
  begin // GL Analtyque par Général ou GL Général par Analytique
    lStTemp := ConvertitCaractereJokers( Y_GENERAL, Y_GENERAL_, 'Y_GENERAL');
    if Trim(lStTemp) <> '' then
      lStWhereCommun := lStWhereCommun + ' AND ' + lStTemp;
  end;

  // Mode Multi-société
  if EstMultiSoc then
  begin
    try
      RequeteMultiDossier( Pages, lStRgrp );
      lQuery := OpenSelect( lStSelect + lStWhere + lStWhereCommun + lStGrpBy, lStRgrp);
      FTobCumulAnoGL.LoadDetailDB( '', '', '', lQuery, False);
    finally
      Ferme(lQuery);
    end;
  end
  else
    FTobCumulAnoGL.LoadDetailFromSql( lStSelect + lStWhere + lStWhereCommun + lStGrpBy );

  // Calcul du Cumul AU
  // GCO - Gestion du Y_QUALIFPIECE
  TraductionTHMultiValComboBox( YQualifPiece, lStValYQualifPiece, lStLibYQualifPiece, 'Y_QUALIFPIECE', False );
  if lStValYQualifPiece <> '' then
    lStWhere := ' AND ' + lStValYQualifPiece;

  lStWhere := lStWhere + ' AND Y_ECRANOUVEAU = "N" AND ' +
              'Y_DATECOMPTABLE >= "' + UsDateTime(StrToDate(DATEDEBUTEXO.Text)) + '" AND ' +
              'Y_DATECOMPTABLE < "' + UsDateTime(StrToDate(YDATECOMPTABLE.Text)) + '"';

  if EstMultiSoc then
  begin
    try
      lQuery := OpenSelect( lStSelect + lStWhere + lStWhereCommun + lStGrpBy, lStRgrp);
      FTobCumulAuGL.LoadDetailDB( '', '', '', lQuery, False);
    finally
      Ferme(lQuery);
    end;
  end
  else
    FTobCumulAuGL.LoadDetailFromSql( lStSelect + lStWhere + lStWhereCommun + lStGrpBy);

  lStTEmp := lStTemp;

end;

////////////////////////////////////////////////////////////////////////////////

Initialization
  registerclasses ( [ TOF_CPGLANA ] ) ;
end.
