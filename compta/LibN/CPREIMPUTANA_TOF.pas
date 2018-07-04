{***********UNITE*************************************************
Auteur  ...... : Stéphane Guillon
Créé le ...... : 09/02/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPREIMPUTANA ()
Suite......... : SG6 09.02.05
Suite......... : Nouvelle fiche pour remplacer ImputAna.pas
Suite......... : Gestion du mode analytique axes croisés
Suite......... : Gestion axes croisés et indépendants en Web 30/11/2006
Mots clefs ... : TOF;CPREIMPUTANA

*****************************************************************}
unit CPREIMPUTANA_TOF;

interface

uses StdCtrls,
  Controls,
  Classes,
  Ent1,
  Saisutil, //PasCreerDate
  HTB97, //TToolBarButton97
  Ventil, // ParamVentil
  CPVENTILTYPECROIS_TOF, //ParamVentilCroisaxe
  Hstatus,
  utob,
  UtilSais, //MajSoldeSectionTOB
  SaisODA, //TrouveEtLanceSaisieODA
  LettUtil, //DateCorrecte
  DelVisue, //VisuPiecesGenere
  UiUtil, //CloseInsidePanel
  SaisComm, //ExecReqMAJNew
  ParamSoc, //GetParamSocSecur
  Ulibanalytique,  //RecherchePremDerAxeVentil
  {$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  mul,
  fe_main, //AglLanceFiche
  Hdb, //THDBGrid
  {$ELSE}
  maineagl, //AglLanceFiche
  eMul,
  {$ENDIF}
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  forms,
  sysutils,
  ed_tools, // VideListe
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF;

type
  TOF_CPREIMPUTANA = class(TOF)
    HM: THMsgBox;
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnDisplay; override;
    procedure OnClose; override;
    procedure OnCancel; override;
  private
    //Controles
    VENTTYPE, Y_EXERCICE, Y_ETABLISSEMENT, Y_AXE, JALGENERE: THValComboBox;
    DATEGENERE, Y_SECTION, Y_DATECOMPTABLE, Y_DATECOMPTABLE_, REFGENERE,
    Y_NUMVENTIL                   : THEdit;
    BZOOMVENTIL, BCHERCHE, BOUVRIR: TToolBarButton97;
    RIMPUT: THRadioGroup;
    RATIO: THNumEdit;
    CONTREP: TCheckBox;
    BChercheClickOriginal: TNotifyEvent;
    {$IFDEF EAGLCLIENT}
    FListe: THGrid;
    {$ELSE}
    FListe: THDBGrid;
    {$ENDIF}

    //Variables
    NumGene, NumVentil: integer;
    ExoGene, OldGene, OldEtab, OldDevise: string;
    OldNumL: integer;
    OldTaux: double;
    DateGene, NowFutur: TDateTime;
    TPieces : TList;
    //Proc
{b fb 14/02/2006}
//    procedure InitComposants; //Initialisation des controles et attribution des événements
    procedure InitComposants(S : string); //Initialisation des controles et attribution des événements
{e fb 14/02/2006}
    procedure InitMsgBox; //Initialisation du message box

    //Func
    function GetChamp(Q: TQuery; NomChamp: string): variant;
    function ChargementVentilType: TOB;

    //Evts
    procedure Y_EXERCICEChange(Sender: TObject);
    procedure RIMPUTClick(Sender: TObject);
    procedure BZOOMVentilClick(Sender: TObject);
    procedure Y_AXEChange(Sender: TOBject);
    procedure BChercheClick(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject);
    procedure FlisteDblClick(Sender: TObject);
    procedure DATEGENEREExit(Sender : TObject);
    procedure BListePieceClick(Sender : TObject);


    //Traitement
    function Traitement(vTobVentil: TOB; Q: TQuery = nil):boolean;
    procedure GenereOD(vTobVentil: TOB; Q: TQuery = nil);
    procedure ModifAna(vTobVentil: TOB; Q: TQuery = nil);
    procedure ContrePasse(vTobAna: TOB ; DEV : RDevise);
    procedure MajMontants(ligneEnCours : integer; vToborigine, vTobModif, vTobVentil : TOB;
      var TotP: double; var TotD: double; DEV: RDevise);
    procedure TraiteQte(vTobModif : TOB ; vTobOrigine : TOB ; ligneEncours : integer ; vTobVentil : TOB ;
      var TotalQte1, TotalQte2, TotalTauxQte1, TotalTauxQte2 : Double );
    procedure ODDevise(vTobOD: TOB);

    //Controle
    function LaDateOk: boolean;
    function JournalOk : boolean ;
    function VerifModelRestriction(Q: TQuery): Boolean; {FP 29/12/2005}
    procedure SlctAllClick (Sender : TObject);
  end;

  //procedure Reimputation Ana;
procedure CC_LanceFicheReimpAna;

implementation

uses
  {$IFDEF MODENT1}
  CPProcGen,
  CPProcMetier,
  CPTypeCons,
  ULibExercice,
  {$ENDIF MODENT1}
  UtilPGI;

procedure CC_LanceFicheReimpAna;
begin
  if PasCreerDate(V_PGI.DateEntree) then Exit;
  if _Blocage(['nrCloture','nrBatch','nrSaisieModif'],False,'nrBatch') then Exit ;
  AGLLanceFiche('CP', 'CPREIMPUTANA', '', '', '');  
end;

procedure TOF_CPREIMPUTANA.OnNew;
begin
  inherited;
end;

procedure TOF_CPREIMPUTANA.OnDelete;
begin
  inherited;
end;

procedure TOF_CPREIMPUTANA.OnUpdate;
begin
  inherited;
end;

procedure TOF_CPREIMPUTANA.OnLoad;
begin
  inherited;
end;

procedure TOF_CPREIMPUTANA.OnArgument(S: string);
begin
  inherited;
  TPieces := TList.Create;

  //Init des controles de la fiche
{b fb 14/02/2006}
//  InitComposants;
  InitComposants(S);
{e fb 14/02/2006}
end;

procedure TOF_CPREIMPUTANA.OnClose;
begin
  inherited;
  _Bloqueur('nrBatch', False);
  HM.Free;
  VideListe(TPieces);
  TPieces.Free;
end;

procedure TOF_CPREIMPUTANA.OnDisplay();
begin
  inherited;
end;

procedure TOF_CPREIMPUTANA.OnCancel();
begin
  inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stephane Guillon
Créé le ...... : 09/02/2005
Modifié le ... :   /  /
Description .. : Procédure qui permet d'inialiser les contrôles et de gérer les
Suite ........ : évenements associé
Mots clefs ... :
*****************************************************************}
{b fb 14/02/2006}
//procedure TOF_CPREIMPUTANA.InitComposants;
procedure TOF_CPREIMPUTANA.InitComposants(S :string);
{e fb 14/02/2006}
var
  i : integer;
  BLISTEPIECE : TToolBarButton97;
  lSt : string; //fb 14/02/2006
begin
  //Init des composants

  Y_NUMVENTIL := THEdit(getcontrol('Y_NUMVENTIL', true));
  Y_AXE := THValComboBox(getcontrol('Y_AXE', true));
  JALGENERE := THValComboBox(getcontrol('JALGENERE', true));
  Y_ETABLISSEMENT := THValComboBox(getcontrol('Y_ETABLISSEMENT', true));
  Y_EXERCICE := THValComboBox(getcontrol('Y_EXERCICE', true));
  RIMPUT := THRadioGroup(getcontrol('RIMPUT', true));
  BZOOMVENTIL := TToolBarButton97(getcontrol('BZOOMVENTIL', true));
  Y_DATECOMPTABLE := THEdit(getcontrol('Y_DATECOMPTABLE', true));
  Y_DATECOMPTABLE_ := THEdit(getcontrol('Y_DATECOMPTABLE_', true));
  Y_SECTION := THEdit(getcontrol('Y_SECTION', true));
  BCHERCHE := TToolBarButton97(getcontrol('BCHERCHE', true));
  BOUVRIR := TToolBarButton97(getcontrol('BOUVRIR', true));
  VENTTYPE := THValComboBox(getcontrol('VENTTYPE', true));
  REFGENERE := THEdit(getcontrol('REFGENERE', true));
  RATIO := THNumEdit(getcontrol('RATIO', true));
  CONTREP := TCheckBox(getcontrol('CONTREP', true));
  DATEGENERE := THEdit(getcontrol('DATEGENERE', true));
  BLISTEPIECE := TToolBarButton97(getcontrol('BListePiece', true));
  {$IFDEF EAGLCLIENT}
  FListe := THGrid(getcontrol('FLISTE', true));
  {$ELSE}
  FListe := THDBGrid(getcontrol('FLISTE', true));
  {$ENDIF}

  //Commun
  if not VH^.MontantNegatif then
  begin
    SetControlEnabled('ContreP', False);
    SetControlChecked('ContreP', False);
  end;

  if VH^.CPExoRef.Code <> '' then
  begin
    Y_EXERCICE.Value := VH^.CPExoRef.Code;
    Y_DATECOMPTABLE.Text := DateToStr(VH^.CPExoRef.Deb);
    Y_DATECOMPTABLE_.Text := DateToStr(VH^.CPExoRef.Fin);
  end
  else
  begin
    Y_EXERCICE.Value := VH^.Entree.Code;
    Y_DATECOMPTABLE.Text := DateToStr(V_PGI.DateEntree);
    Y_DATECOMPTABLE_.Text := DateToStr(V_PGI.DateEntree);
  end;

  //MsgBox
  HM := THMsgBox.Create(Ecran);
  InitMsgBox;

  SetControlText('Y_DEVISE', V_PGI.DevisePivot);
  SetControlText('DATEGENERE', DateToStr(V_PGI.DateEntree));
  if JALGENERE.Values.Count > 0 then JALGENERE.Value := JALGENERE.Values[0];
  PositionneEtabUser(Y_ETABLISSEMENT);
  SetControlProperty('PAVANCE', 'Tabvisible', False);

  //Evenements
  Y_EXERCICE.OnChange := Y_EXERCICEChange;
  RIMPUT.OnClick := RIMPUTClick;
  BZOOMVENTIL.OnClick := BZOOMVENTILClick;
  BChercheClickOriginal := BCHERCHE.OnClick;
  BCHERCHE.OnClick := BChercheClick;
  BOUVRIR.OnClick := BOuvrirClick;
  FLISTE.OnDblClick := FlisteDblClick;
  DATEGENERE.OnExit := DATEGENEREExit;
  BLISTEPIECE.OnClick := BListePieceClick;

  {FQ22310 18.02.08 YMO}
  TFMul(Ecran).bSelectAll.Visible := True;
  TFMul(Ecran).bSelectAll.OnClick := SlctAllClick;

  //Mode Analytique classique
  if not VH^.AnaCroisaxe then
  begin
    SetControlProperty('PSection', 'Tabvisible', False);
    for i := 1 to 5 do
    begin
      SetControlText('Y_SOUSPLAN' + IntToStr(i), '');
    end;
    Y_AXE.Value := Y_AXE.Values[0];
    SetControlText('Y_SECTION', VH^.Cpta[AxeToFb(Y_Axe.Value)].Attente);

    Y_AXE.OnChange := Y_AXEChange;
    {b Thl 30/11/2006 on prend les Y_NUMVENTIL >=1}
    Y_NUMVENTIL.Operateur := Superieur;
    {e Thl}

  end
  else
  begin
    Y_AXE.Visible := False;
    Y_AXE.Text := '';
    Y_SECTION.Text := '';
    Y_SECTION.Visible := False;
    SetControlVisible('TY_SECTION', False);
    SetControlVisible('TY_AXE', False);

    for i := 1 to MaxAxe do
    begin
      SetControlVisible('Y_SOUSPLAN' + IntToStr(i), GetParamSocSecur('SO_VENTILA' + IntToStr(i),False));
      SetControlVisible('TY_SOUSPLAN' + IntToStr(i), GetParamSocSecur('SO_VENTILA' + IntToStr(i),False));
    end;
  end;

{b fb 14/02/2006}
  lSt := ReadTokenSt(S);
  if lSt <> '' then
    Y_Axe.Text := lSt;

  lSt := ReadTokenSt(S);
  if lSt <> '' then
    SetControlText('Y_GENERAL', lSt);

  lSt := ReadTokenSt(S);
  if lSt <> '' then
    Y_EXERCICE.Value:=lSt;

  lSt := ReadTokenSt(S);
  if lSt <> '' then
    SetControlText('Y_SECTION', lSt);

  lSt := ReadTokenSt(S);
  if lSt <> '' then
    SetControlText('Y_JOURNAL', lSt);

  lSt := ReadTokenSt(S);
  if lSt <> '' then
    SetControlText('Y_DEVISE', lSt);

  lSt := ReadTokenSt(S);
  if lSt <> '' then
    SetControlText('Y_QUALIFPIECE', lSt);

  lSt := ReadTokenSt(S);
  if lSt <> '' then
    SetControlText('Y_NUMEROPIECE', lSt);
{e fb 14/02/2006}

end;

procedure TOF_CPREIMPUTANA.Y_EXERCICEChange(Sender: TObject);
begin
  ExoToDates(Y_EXERCICE.Value, Y_DATECOMPTABLE, Y_DATECOMPTABLE_);
end;

procedure TOF_CPREIMPUTANA.RIMPUTClick(Sender: TObject);
begin
  SetControlProperty('RATIO', 'Value', 1);
  if RIMPUT.ItemIndex = 0 then
  begin
    SetControlEnabled('DATEGENERE', False);
    SetControlEnabled('RATIO', False);
    SetControlEnabled('CONTREP', False);
    SetControlEnabled('JALGENERE', False);
    SetControlEnabled('REFGENERE', False);
  end
  else
  begin
    SetControlEnabled('DATEGENERE', True);
    SetControlEnabled('RATIO', True);
    SetControlEnabled('CONTREP', True);
    SetControlEnabled('JALGENERE', True);
    SetControlEnabled('REFGENERE', True);
  end;
end;

procedure TOF_CPREIMPUTANA.BZOOMVentilClick(Sender: TObject);
begin
  if VENTTYPE.Text = '' then Exit;
  if VH^.AnaCroisaxe then ParamVentilCroise('TY', VENTTYPE.Value, taConsult, True)
  else ParamVentil('TY', VENTTYPE.Value, '12345', taConsult, True);
end;

procedure TOF_CPREIMPUTANA.Y_AXEChange(Sender: TOBject);
var
  SNomTablette: string;
begin
  if Y_AXE.Value = 'A1' then
  begin
    SNomTablette := 'TZSECTION';
  end
  else if Y_AXE.Value = 'A2' then
  begin
    SNomTablette := 'TZSECTION2';
  end
  else if Y_AXE.Value = 'A3' then
  begin
    SNomTablette := 'TZSECTION3';
  end
  else if Y_AXE.Value = 'A4' then
  begin
    SNomTablette := 'TZSECTION4';
  end
  else if Y_AXE.Value = 'A4' then
  begin
    SNomTablette := 'TZSECTION5';
  end;

  SetControlProperty('Y_SECTION', 'Datatype', SNomTablette);
  SetControlText('Y_SECTION', '');
end;

procedure TOF_CPREIMPUTANA.InitMsgBox; //Initialisation du message box
begin
  HM.Mess.Add('0;' + Ecran.Caption + ';Vous devez renseigner une section analytique;W;O;O;O;');
  HM.Mess.Add('1;' + Ecran.Caption + ';Vous devez renseigner une grille de répartition;W;O;O;O;');
  HM.Mess.Add('2;' + Ecran.Caption + ';Vous n''avez sélectionné aucune écriture analytique;W;O;O;O;');
  HM.Mess.Add('3;' + Ecran.Caption + ';La grille de répartition n''est pas correcte;W;O;O;O;');
  HM.Mess.Add('4;' + Ecran.Caption + ';Attention, seules les ventilations à 100% sur la section à ré-imputer seront reventilées sur la grille;E;O;O;O;');
  HM.Mess.Add('5;' + Ecran.Caption + ';Vous devez renseigner une date valide;W;O;O;O;');
  HM.Mess.Add('6;' + Ecran.Caption + ';Le journal d''OD analytique est fermé;W;O;O;O;');
  HM.Mess.Add('7;' + Ecran.Caption + ';Le journal d''OD analytique ne possède pas de facturier;W;O;O;O;');
  HM.Mess.Add('8;' + Ecran.Caption + ';Le journal d''OD analytique n''est pas sur le même axe;W;O;O;O;');
  HM.Mess.Add('9;' + Ecran.Caption + ';Le traitement s''est correctement effectué;I;O;O;O;');
  HM.Mess.Add('10;' + Ecran.Caption + ';Le traitement s''est correctement effectué. Voulez-vous visualiser les écritures générées ?;Q;YN;Y;Y;');
  HM.Mess.Add('11;' + Ecran.Caption + ';Confirmez-vous le traitement de ré-imputation analytique ?;Q;YNC;Y;Y;');
  HM.Mess.Add('12;' + Ecran.Caption + ';Attention ! Aucune réimputation n''a été générée;W;O;O;O;');
  {b FP 29/12/2005}
  HM.Mess.Add('13;' + Ecran.Caption + ';Il existe des incompatibilités entre le modèle de ventilations à appliquer et les modèles de restrictions analytiques attachés au comptes généraux sélectionnés.;W;O;O;O;');  {FP 19/04/2006}
  {e FP 29/12/2005}
end;

procedure TOF_CPREIMPUTANA.BChercheClick(Sender: TOBject);
begin
  if not (VH^.AnaCroisaxe) and (Y_SECTION.Text = '') then
  begin
    HM.Execute(0, '', '');
  end
  else
  begin
    BChercheClickOriginal(Sender);
  end;
end;

procedure TOF_CPREIMPUTANA.BOuvrirClick(Sender: TObject);
var
  Q: TQuery;
  vTobVentil: TOB;
  i: integer;
  bLigneReimput : boolean;
begin
  bLigneReimput := false;

  if not (VH^.AnaCroisaxe) and (VENTTYPE.Text = '') then
  begin
    HM.Execute(1, '', '');
    Exit;
  end;
  if not (FListe.AllSelected) and (FListe.nbSelected <= 0) then
  begin
    HM.Execute(2, '', '');
    Exit;
  end;

  //Vérif date et journaux si en Gen d'OD
  if RIMPUT.ItemIndex = 1 then
  begin
    if not LaDateOk then Exit;
    if (not VH^.AnaCroisaxe) and (not JournalOk) then Exit;
  end;

  if HM.Execute(11,'','') <> mrYes then Exit;

  //Avertissement prise en compte des ligne à 100%
  if (not VH^.AnaCroisaxe) and (RIMPUT.ItemIndex <> 1) then  //thl ajoute la condition
    HM.Execute(4, '', '');


  {$IFDEF EAGLCLIENT}
  Q := TFMul(Ecran).Q.TQ;
  {$ELSE}
  Q := TFMul(Ecran).Q;
  {$ENDIF}

  vTobVentil := ChargementVentilType;
  if vTobVentil = nil then
  begin
    HM.Execute(3, '', '');
    Exit;
  end;

  //Init variables
  OldGene := '';
  OldEtab := '';
  OldDevise := '';
  OldNumL := -1;
  OldTaux := 0.0;
  NumVentil := 0;
  ExoGene := QuelExoDT(DateGene);
  NowFutur := NowH;

  VideListe(TPieces);
  {b FP 29/12/2005}
  if not VerifModelRestriction(Q) then
  begin
    HM.Execute(13,'','') ;
  end
  else
  begin
    Q.First;
   {e FP 29/12/2005}
    if FListe.AllSelected then InitMove(Q.RecordCount, Ecran.Caption)
                          else InitMove(FListe.NbSelected, Ecran.Caption);
    try
      BeginTrans;
      try
        if FListe.AllSelected then
        begin
          Q.First;
          while not Q.eof do
          begin
            MoveCur(False);
            bLigneReimput := Traitement(vTobVentil, Q) or bLigneReimput;
            Q.Next;
          end;
        end
        else
        begin
          for i := 0 to FListe.NbSelected - 1 do
          begin
            MoveCur(False);
            FListe.GotoLeBookmark(i);
            {b Thl 30/11/2006 }
            {$IFDEF EAGLCLIENT}
            Q.Seek(FListe.Row-1);
            {$ENDIF}
            {e Thl }
            bLigneReimput := Traitement(vTobVentil, nil) or bLigneReimput;
          end;
        end;
        Committrans;

      except
        on e: exception do
        begin
          rollback;
          PGIInfo(e.message);
          FreeAndNil(vTobVentil);
          Exit;
        end;
      end;
    finally
      FiniMove;
    end; 
  end;

  if not bLigneReimput then HM.Execute(12, '', '');

  FreeAndNil(vTobVentil);

  //Fin
  if RIMPUT.ItemIndex = 0 then begin
    {JP 18/05/05 : pour éviter le message si aucune ligne n'a été traitée}
    if bLigneReimput then
      HM.Execute(9,'','');
  end
  else
  begin
    if (bLigneReimput) and (HM.Execute(10,'','') = mrYes) then
    begin
      VisuPiecesGenere(TPieces, EcrAna,12);
    end;
  end;

  //Maj du grid
  TFMul(Ecran).BChercheClick(nil);

end;

procedure TOF_CPREIMPUTANA.GenereOD(vTobVentil: TOB; Q: TQuery = nil);
var
  vTobAna: TOB;
  vTobOD: TOB;
  rupture: boolean;
  i: integer;
  TotP, TotD, TotalQte1, TotalQte2, TotalTauxQte1, TotalTauxQte2: double;
  DEV: RDevise;
begin
  rupture := False;

  //Création de la tob analytique
  vTobAna := TOB.Create('ANALYTIQ', nil, -1);
  try
    if Q <> nil then
    begin
      {JP 17/05/05 : A utiliser avec parcimonie, car SelectDB ne fait pas un select * mais un
                     select Champs du Query
      vTobAna.SelectDB('', Q);}
      vTobAna.SetString('Y_JOURNAL', Q.FindField('Y_JOURNAL').AsString);
      vTobAna.SetString('Y_EXERCICE', Q.FindField('Y_EXERCICE').AsString);
      vTobAna.SetDateTime('Y_DATECOMPTABLE', Q.FindField('Y_DATECOMPTABLE').AsDateTime);
      vTobAna.SetInteger('Y_NUMEROPIECE', Q.FindField('Y_NUMEROPIECE').AsInteger);
      vTobAna.SetInteger('Y_NUMLIGNE', Q.FindField('Y_NUMLIGNE').AsInteger);
      vTobAna.SetString('Y_AXE', Q.FindField('Y_AXE').AsString);
      vTobAna.SetInteger('Y_NUMVENTIL', Q.FindField('Y_NUMVENTIL').AsInteger);
      vTobAna.SetString('Y_QUALIFPIECE', Q.FindField('Y_QUALIFPIECE').AsString);
    end
    else
    begin
      vTobAna.PutValue('Y_JOURNAL', GetField('Y_JOURNAL'));
      vTobAna.PutValue('Y_EXERCICE', GetField('Y_EXERCICE'));
      vTobAna.PutValue('Y_DATECOMPTABLE', GetField('Y_DATECOMPTABLE'));
      vTobAna.PutValue('Y_NUMEROPIECE', GetField('Y_NUMEROPIECE'));
      vTobAna.PutValue('Y_NUMLIGNE', GetField('Y_NUMLIGNE'));
      vTobAna.PutValue('Y_AXE', GetField('Y_AXE'));
      vTobAna.PutValue('Y_NUMVENTIL', GetField('Y_NUMVENTIL'));
      vTobAna.PutValue('Y_QUALIFPIECE', GetField('Y_QUALIFPIECE'));
    end;
    vTobAna.LoadDB;

    //Gestion de la rupture
    if ((OldGene <> vTobAna.GetValue('Y_GENERAL')) or
      (OldEtab <> vTobAna.GetValue('Y_ETABLISSEMENT')) or
      (OldDevise <> vTobAna.GetValue('Y_DEVISE')) or
      (OldNumL <> vTobAna.GetValue('Y_NUMLIGNE')) or
      (OldTaux <> vTobAna.GetValue('Y_TAUXDEV'))) then Rupture := True;

    OldGene := vTobAna.GetValue('Y_GENERAL');
    OldEtab := vTobAna.GetValue('Y_ETABLISSEMENT');
    OldDevise := vTobAna.GetValue('Y_DEVISE');
    OldTaux := vTobAna.GetValue('Y_TAUXDEV');
    OldNumL := vTobAna.GetValue('Y_NUMLIGNE');

    if rupture then
    begin
      NumGene := GetNewNumJal(JALGENERE.Value, True, DateGene);
      if NumGene <= 0 then V_PGI.IoError := oeUnknown;
      NumVentil := 0;
    end;

    TotD := 0.0;
    TotP := 0.0;
    TotalQte1 := 0.0;
    TotalQte2 := 0.0;
    TotalTauxQte1 := 0.0;
    TotalTauxQte2 := 0.0;

    //Récuperation devise
    DEV.Code := vTobAna.GetValue('Y_DEVISE');
    GetInfosDevise(DEV);

    //Génération OD à partir de la ventilation
    for i := 0 to vTobVentil.Detail.Count - 1 do
    begin
      vTobOD := TOB.Create('ANALYTIQ', nil, -1);
      try
        vTobOD.PutValue('Y_GENERAL', vTobAna.GetValue('Y_GENERAL'));
        vTobOD.PutValue('Y_ETABLISSEMENT', vTobAna.GetValue('Y_ETABLISSEMENT'));
        vTobOD.PutValue('Y_LIBELLE', vTobAna.GetValue('Y_LIBELLE'));
        vTobOD.PutValue('Y_TYPEANOUVEAU', vTobAna.GetValue('Y_TYPEANOUVEAU'));
        vTobOD.PutValue('Y_DEVISE', vTobAna.GetValue('Y_DEVISE')); // ?
        vTobOD.PutValue('Y_TAUXDEV', vTobAna.GetValue('Y_TAUXDEV')); //?
        vTobOD.PutValue('Y_AXE', vTobAna.GetValue('Y_AXE'));
        vTobOD.PutValue('Y_SECTION', vTobVentil.Detail[i].GetValue('V_SECTION'));
        vTobOD.PutValue('Y_QUALIFPIECE', 'N');
        vTobOD.PutValue('Y_ECRANOUVEAU', 'N');
        vTobOD.PutValue('Y_CREERPAR', 'GEN');
        vTobOD.PutValue('Y_DATEMODIF', NowFutur);

        if VH^.AnaCroisaxe then
        begin
          vTobOD.PutValue('Y_SOUSPLAN1', vTobVentil.Detail[i].GetValue('V_SOUSPLAN1'));
          vTobOD.PutValue('Y_SOUSPLAN2', vTobVentil.Detail[i].GetValue('V_SOUSPLAN2'));
          vTobOD.PutValue('Y_SOUSPLAN3', vTobVentil.Detail[i].GetValue('V_SOUSPLAN3'));
          vTobOD.PutValue('Y_SOUSPLAN4', vTobVentil.Detail[i].GetValue('V_SOUSPLAN4'));
          vTobOD.PutValue('Y_SOUSPLAN5', vTobVentil.Detail[i].GetValue('V_SOUSPLAN5'));
        end;

        Inc(NumVentil);
        if NumVentil = 1 - i then vTobOD.PutValue('Y_TYPEMVT', 'AE')
                             else vTobOD.PutValue('Y_TYPEMVT', 'AL');
        vTobOD.PutValue('Y_DATECOMPTABLE', DateGene);
        vTobOD.PutValue('Y_NUMEROPIECE', NumGene);
        vTobOD.PutValue('Y_PERIODE', GetPeriode(DateGene));
        vTobOD.PutValue('Y_SEMAINE', NumSemaine(DateGene));
        vTobOD.PutValue('Y_NUMLIGNE', 0);
        vTobOD.PutValue('Y_EXERCICE', ExoGene);
        vTobOD.PutValue('Y_NATUREPIECE', 'OD');
        vTobOD.PutValue('Y_TYPEANALYTIQUE', 'X');
        vTobOD.PutValue('Y_TOTALECRITURE', 0.0);
        vTobOD.PutValue('Y_TOTALDEVISE', 0.0);
        vTobOD.PutValue('Y_JOURNAL', JALGENERE.Value);
        vTobOD.PutValue('Y_POURCENTAGE', 0.0);

        if REFGENERE.Text <> '' then vTobOD.PutValue('Y_REFINTERNE', REFGENERE.Text)
        else vTobOD.PutValue('Y_REFINTERNE', vTobAna.GetValue('Y_REFINTERNE'));

        {26.06.07 FQ18513 YMO}
        vTobOD.PutValue('Y_DATEREFEXTERNE', vTobAna.GetValue('Y_DATEREFEXTERNE'));

        vTobOD.PutValue('Y_NUMVENTIL', NumVentil);

        //Maj Montants
        MajMontants(i, vTobAna, vTobOD, vTobVentil, TotP, TotD, DEV);

        //Maj Quantite
        TraiteQte(vTobOD, vTobAna, i, vTobVentil, TotalQte1, TotalQte2, TotalTauxQte1, TotalTauxQte2);

        //Maj devise
        ODDevise(vTobOD);

        //Insertion ds TList pour visu
        if NumVentil = 1 then TPieces.Add(vTobOD);

        //Insertion ds la base
        vTobOD.InsertDB(nil);

        //Maj Solde
        MajSoldeSectionTOB(vTobOD, True);

        //Maj journaux
        MajJournalAnaTob(vTobOD, False, True);
      finally
        {La tob correspondant au NumVentil = 1 sera vidée à la fermeture la fiche lors de la
         destruction de la liste TPieces}
        if NumVentil <> 1 then FreeAndNil(vTobOD);
      end;
    end;
    ContrePasse(vTobAna, DEV);
  finally
    FreeAndNil(vTobAna);
  end;
end;

procedure TOF_CPREIMPUTANA.ModifAna(vTobVentil: TOB; Q: TQuery = nil);
var
  vTobAna, vTobOrigine: TOB;
  DEV: RDevise;
  TotP, TotD, TotalQte1, TotalQte2, TotalTauxQte1, TotalTauxQte2 : double;
  i: integer;
begin
  //Création de la tob analytique
  vTobOrigine := TOB.Create('ANALYTIQ', nil, -1);
  if Q <> nil then
  begin
    {JP 17/05/05 : A utiliser avec parcimonie, car SelectDB ne fait pas un select * mais un
                   select Champs du Query
    vTobOrigine.SelectDB('', Q);}
    vTobOrigine.SetString('Y_JOURNAL', Q.FindField('Y_JOURNAL').AsString);
    vTobOrigine.SetString('Y_EXERCICE', Q.FindField('Y_EXERCICE').AsString);
    vTobOrigine.SetDateTime('Y_DATECOMPTABLE', Q.FindField('Y_DATECOMPTABLE').AsDateTime);
    vTobOrigine.SetInteger('Y_NUMEROPIECE', Q.FindField('Y_NUMEROPIECE').AsInteger);
    vTobOrigine.SetInteger('Y_NUMLIGNE', Q.FindField('Y_NUMLIGNE').AsInteger);
    vTobOrigine.SetString('Y_AXE', Q.FindField('Y_AXE').AsString);
    vTobOrigine.SetInteger('Y_NUMVENTIL', Q.FindField('Y_NUMVENTIL').AsInteger);
    vTobOrigine.SetString('Y_QUALIFPIECE', Q.FindField('Y_QUALIFPIECE').AsString);
  end
  else
  begin
    vTobOrigine.PutValue('Y_JOURNAL', GetField('Y_JOURNAL'));
    vTobOrigine.PutValue('Y_EXERCICE', GetField('Y_EXERCICE'));
    vTobOrigine.PutValue('Y_DATECOMPTABLE', GetField('Y_DATECOMPTABLE'));
    vTobOrigine.PutValue('Y_NUMEROPIECE', GetField('Y_NUMEROPIECE'));
    vTobOrigine.PutValue('Y_NUMLIGNE', GetField('Y_NUMLIGNE'));
    vTobOrigine.PutValue('Y_AXE', GetField('Y_AXE'));
    vTobOrigine.PutValue('Y_NUMVENTIL', GetField('Y_NUMVENTIL'));
    vTobOrigine.PutValue('Y_QUALIFPIECE', GetField('Y_QUALIFPIECE'));
  end;
  vTobOrigine.LoadDB;

  //Mise a jour solde section
  MajSoldeSectionTOB(vTobOrigine, False);

  //On efface la ligne d'analytique
  vTobOrigine.DeleteDB;

  //Creation tob de travail
  vTobAna := TOB.Create('ANALYTIQ', nil, -1);
  vTobAna.Dupliquer(vTobOrigine, true, true);

  //Init des totaux
  TotP := 0.00;
  TotD := 0.00;
  TotalQte1 := 0.0;
  TotalQte2 := 0.0;
  TotalTauxQte1 := 0.0;
  TotalTauxQte2 := 0.0;

  //Récupération infos de la devise
  DEV.Code := vTobAna.GetValue('Y_DEVISE');
  GetInfosDevise(DEV);

  //Reventilation
  for i := 0 to vTobVentil.Detail.Count - 1 do
  begin

    //Mise à jour des champs nécessaires
    vTobAna.PutValue('Y_NUMVENTIL', vTobVentil.Detail[i].GetInteger('V_NUMEROVENTIL'));
    vTobAna.PutValue('Y_SECTION', vTobVentil.Detail[i].GetString('V_SECTION'));
    vTobAna.PutValue('Y_POURCENTAGE', vTobVentil.Detail[i].GetDouble('V_TAUXMONTANT'));
    vTobAna.PutValue('Y_POURCENTQTE1', vTobVentil.Detail[i].GetDouble('V_TAUXQTE1'));
    vTobAna.PutValue('Y_POURCENTQTE2', vTobVentil.Detail[i].GetDouble('V_TAUXQTE2'));

    //Mode Croisaxe
    if VH^.AnaCroisaxe then
    begin
      vTobAna.PutValue('Y_SOUSPLAN1', vTobVentil.Detail[i].GetValue('V_SOUSPLAN1'));
      vTobAna.PutValue('Y_SOUSPLAN2', vTobVentil.Detail[i].GetValue('V_SOUSPLAN2'));
      vTobAna.PutValue('Y_SOUSPLAN3', vTobVentil.Detail[i].GetValue('V_SOUSPLAN3'));
      vTobAna.PutValue('Y_SOUSPLAN4', vTobVentil.Detail[i].GetValue('V_SOUSPLAN4'));
      vTobAna.PutValue('Y_SOUSPLAN5', vTobVentil.Detail[i].GetValue('V_SOUSPLAN5'));
    end;

    //Maj montants
    MajMontants(i, vTobOrigine, vTobAna, vTobVentil, TotP, TotD, DEV);

    //Maj quantite
    TraiteQte(vTobAna, vTobOrigine, i, vTobVentil, TotalQte1, TotalQte2, TotalTauxQte1, TotalTauxQte2);


    //Insertion
    vTobAna.InsertDB(nil);

    //Mise à jour des solde section
    MajSoldeSectionTOB(vTobAna, True);
  end;

  FreeAndNil(vTobAna);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stephane Guillon
Créé le ...... : 11/02/2005
Modifié le ... :   /  /
Description .. : Fonction qui retourne la valeur de champ de la query
Mots clefs ... : RETOURNE VALEUR QUERY
*****************************************************************}
function TOF_CPREIMPUTANA.GetChamp(Q: TQuery; NomChamp: string): variant;
begin
  if NomChamp = '' then
  begin
    result := null;
    Exit;
  end;

  if Q = nil then
  begin
    result := GetField(NomChamp);
  end
  else
  begin
    result := Q.FindField(NomChamp).AsVariant;
  end;
end;

function TOF_CPREIMPUTANA.Traitement(vTobVentil: TOB; Q: TQuery = nil):boolean;
begin
  result := False;
  //Si ventilation pas a 100%
  if VH^.AnaCroisaxe then begin // Thl  30/11/2006 ajoute le test.
    if Double(GetChamp(Q, 'Y_POURCENTAGE')) <> 100.0 then Exit
    end
  {b Thl 30/11/2006}
  else
  if not VH^.AnaCroisaxe  and (RIMPUT.ItemIndex = 0) then begin
     if Double(GetChamp(Q, 'Y_POURCENTAGE')) <> 100.0 then Exit;
     if Integer(GetChamp(Q, 'Y_NUMVENTIL')) <> 1 then Exit ;
     end;
  {e Thl }
  result := True;
  if RIMPUT.ItemIndex = 0 then
  begin
    ModifAna(vTobVentil, Q);
  end
  else
  begin
    GenereOD(vTobVentil, Q);
  end;
end;

function TOF_CPREIMPUTANA.ChargementVentilType: TOB;
var
  vTobVentil: TOB;
  vPourcent: double;
  i: integer;
  ParamAxe : TParamAxe;
  sAxe : string;
begin
  ParamAxe := RecherchePremDerAxeVentil;

  if ParamAxe.premier_axe = 0 then sAxe := Y_AXE.Value[2]
  else sAxe := IntToStr(ParamAxe.premier_axe);

  vPourcent := 0.0;

  vTobVentil := TOB.Create('$VENTIL', nil, -1);
  vTobVentil.LoadDetailDBFromSQL('VENTIL', 'SELECT * FROM VENTIL WHERE '
    + 'V_NATURE="TY' + sAxe + '" AND '
    + 'V_COMPTE="' + VENTTYPE.Value + '" ORDER BY V_NUMEROVENTIL');

  if vTobVentil.Detail.Count = 0 then FreeAndNil(vTobVentil);

  if vTobVentil <> nil then
  begin
    //Calcul du total %
    for i := 0 to vTobVentil.Detail.Count - 1 do
    begin
      vPourcent := vPourcent + vTobVentil.Detail[i].GetDouble('V_TAUXMONTANT');
    end;

    //Vérification pourcentage à 100
    if Arrondi(vPourcent - 100.0, V_PGI.OkDecP) <> 0 then
    begin
      FreeAndNil(vTobVentil);
    end;
  end;

  result := vTobVentil;
end;

procedure TOF_CPREIMPUTANA.FlisteDblClick(Sender: TOBject);
var
  Q: TQuery;                                                                                             
begin
  {$IFDEF EAGLCLIENT}
  Q := TFMul(Ecran).Q.TQ;
  {b Thl 30/11/2006 }
  Q.Seek(FListe.Row-1);
  {e Thl }

  {$ELSE}
  Q := TFMul(Ecran).Q;
  {$ENDIF}

  if ((Q.Eof) and (Q.Bof)) then Exit;  
  TrouveEtLanceSaisieODA(Q, taConsult);
end;

procedure TOF_CPREIMPUTANA.MajMontants(ligneEnCours : integer; vToborigine, vTobModif, vTobVentil : TOB; var TotP: double;
  var TotD: double; DEV: RDevise);
var
  Pourc, dratio : double;
  ligneMax : integer;
begin
  Pourc := vTobVentil.Detail[ligneEnCours].GetValue('V_TAUXMONTANT');
  ligneMax := vTobVentil.Detail.Count - 1;

  if RIMPUT.ItemIndex = 0 then dRatio := 1.0 else dRatio := RATIO.Value;
  if ligneEnCours < ligneMax then
  begin
    vTobModif.PutValue('Y_DEBIT', Arrondi(vTobOrigine.GetValue('Y_DEBIT') * PourC * dRatio / 100.0, V_PGI.OkDecV));
    vTobModif.PutValue('Y_CREDIT', Arrondi(vTobOrigine.GetValue('Y_CREDIT') * PourC * dRatio / 100.0, V_PGI.OkDecV));
    vTobModif.PutValue('Y_DEBITDEV', Arrondi(vTobOrigine.GetValue('Y_DEBITDEV') * PourC * dRatio / 100.0, DEV.Decimale));
    vTobModif.PutValue('Y_CREDITDEV', Arrondi(vTobOrigine.GetValue('Y_CREDITDEV') * Pourc * dRatio / 100.0, DEV.Decimale));

    TotD := Arrondi(TotD + vTobModif.GetValue('Y_DEBITDEV') + vTobModif.GetValue('Y_CREDITDEV'), DEV.Decimale);
    TotP := Arrondi(TotP + vTobModif.GetValue('Y_DEBIT') + vTobModif.GetValue('Y_CREDIT'), V_PGI.OkDecV);
  end
  else
    //Si derniere ligne
  begin
    {FQ 15865 : erreur de frappe  : les montants de vTobModif sont vides
     if vTobModif.GetValue('Y_DEBITDEV') <> 0 then}
    if vTobOrigine.GetValue('Y_DEBITDEV') <> 0 then
    begin
      vTobModif.PutValue('Y_DEBITDEV', Arrondi((vTobOrigine.GetValue('Y_DEBITDEV') * dRatio) - TotD, DEV.Decimale));
      vTobModif.PutValue('Y_DEBIT', Arrondi((vTobOrigine.GetValue('Y_DEBIT') * dRatio) - TotP, V_PGI.OkDecV));
    end
    else
    begin
      vTobModif.PutValue('Y_CREDITDEV', Arrondi((vTobOrigine.GetValue('Y_CREDITDEV') * dRatio) - TotD, DEV.Decimale));
      vTobModif.PutValue('Y_CREDIT', Arrondi((vTobOrigine.GetValue('Y_CREDIT') * dRatio) - TotP, V_PGI.OkDecV));
    end;
  end;
end;

procedure TOF_CPREIMPUTANA.ContrePasse(vTobAna: TOB ; DEV: RDevise);
var
  vTobOD: TOB;
  dRatio: double;
  sAxe : string;
begin
  if not VH^.AnaCroisaxe then sAxe := Y_AXE.Value
  else sAxe := 'A' + IntToStr(RecherchePremDerAxeVentil.premier_axe);

  dRatio := RATIO.Value;
  Inc(NumVentil);

  vTobOD := TOB.Create('ANALYTIQ', nil, -1);
  vTobOD.PutValue('Y_SECTION', vTobAna.GetValue('Y_SECTION'));
  vTobOD.PutValue('Y_GENERAL', vTobAna.GetValue('Y_GENERAL'));
  vTobOD.PutValue('Y_ETABLISSEMENT', vTobAna.GetValue('Y_ETABLISSEMENT'));
  vTobOD.PutValue('Y_TYPEANOUVEAU', vTobAna.GetValue('Y_TYPEANOUVEAU'));
  vTobOD.PutValue('Y_DEVISE', vTobAna.GetValue('Y_DEVISE'));
  vTobOD.PutValue('Y_TAUXDEV', vTobAna.GetValue('Y_TAUXDEV'));
  vTobOD.PutValue('Y_DATETAUXDEV', vTobAna.GetValue('Y_DATETAUXDEV'));
  vTobOD.PutValue('Y_JOURNAL', JALGENERE.Value);
  vTobOD.PutValue('Y_NUMVENTIL', NumVentil);
  vTobOD.PutValue('Y_REFINTERNE', vTobAna.GetValue('Y_REFINTERNE'));
  vTobOD.PutValue('Y_LIELLE', vTobAna.GetValue('Y_LIBELLE'));
  vTobOD.PutValue('Y_AXE', sAxe);
  vTobOD.PutValue('Y_QUALIFPIECE', 'N');
  vTobOD.PutValue('Y_ECRANOUVEAU', 'N');
  vTobOD.PutValue('Y_POURCENTAGE', 0);
  vTobOD.PutValue('Y_DATEMODIF', NowFutur);
  vTobOD.PutValue('Y_TYPEMVT', 'AL');
  vTobOD.PutValue('Y_DATECOMPTABLE', DateGene);
  vTobOD.PutValue('Y_PERIODE', GetPeriode(DateGene));
  vTobOD.PutValue('Y_SEMAINE', NumSemaine(DateGene));
  vTobOD.PutValue('Y_NUMLIGNE', 0);
  vTobOD.PutValue('Y_NUMEROPIECE', NumGene);
  vTobOD.PutValue('Y_TOTALECRITURE', 0);
  vTobOD.PutValue('Y_NATUREPIECE', 'OD');
  vTobOD.PutValue('Y_TOTALDEVISE', 0);
  vTobOD.PutValue('Y_EXERCICE', ExoGene);
  vTobOD.PutValue('Y_TYPEANALYTIQUE', 'X');

  if CONTREP.Checked then
  begin
    vTobOD.PutValue('Y_DEBIT', Arrondi(-vTobAna.GetValue('Y_DEBIT') * dRatio, V_PGI.OkDecV));
    vTobOD.PutValue('Y_CREDIT', Arrondi(-vTobAna.GetValue('Y_CREDIT') * dRatio, V_PGI.OkDecV));
    vTobOD.PutValue('Y_DEBITDEV', Arrondi(-vTobAna.Getvalue('Y_DEBITDEV') * dRatio, DEV.Decimale));
    vTobOD.PutValue('Y_CREDITDEV', Arrondi(-vTobAna.GetValue('Y_CREDITDEV') * dRatio, DEV.Decimale));
  end
  else
  begin
    vTobOD.PutValue('Y_DEBIT', Arrondi(vTobAna.GetValue('Y_CREDIT') * dRatio, V_PGI.OkDecV));
    vTobOD.PutValue('Y_CREDIT', Arrondi(vTobAna.GetValue('Y_DEBIT') * dRatio, V_PGI.OkDecV));
    vTobOD.PutValue('Y_DEBITDEV', Arrondi(vTobAna.Getvalue('Y_CREDITDEV') * dRatio, DEV.Decimale));
    vTobOD.PutValue('Y_CREDITDEV', Arrondi(vTobAna.GetValue('Y_DEBITDEV') * dRatio, DEV.Decimale));
  end;

  vTobOD.PutValue('Y_QUALIFECRQTE1', vTobAna.GetValue('Y_QUALIFECRQTE1'));
  vTobOD.PutValue('Y_QUALIFECRQTE2', vTobAna.GetValue('Y_QUALIFECRSTE2'));
  vTobOD.PutValue('Y_TOTALQTE1', Arrondi(vTobAna.GetValue('Y_QTE1') * dRatio, 3));
  vTobOD.PutValue('Y_TOTALQTE2', Arrondi(vTobAna.GetValue('Y_QTE2') * dRatio, 3));
  vTobOD.PutValue('Y_POURCENTQTE1', 100.0);
  vTobOD.PutValue('Y_POURCENTQTE2', 100.0);
  vTobOD.PutValue('Y_QTE1', -1 * Arrondi(vTobAna.GetValue('Y_QTE1') * dRatio, 3));
  vTobOD.PutValue('Y_QTE2', -1 * Arrondi(vTobAna.GetValue('Y_QTE2') * dRatio, 3));

  //Mode Croisaxe
  if VH^.AnaCroisaxe then
  begin
    vTobOD.PutValue('Y_SOUSPLAN1', vTobAna.GetValue('Y_SOUSPLAN1'));
    vTobOD.PutValue('Y_SOUSPLAN2', vTobAna.GetValue('Y_SOUSPLAN2'));
    vTobOD.PutValue('Y_SOUSPLAN3', vTobAna.GetValue('Y_SOUSPLAN3'));
    vTobOD.PutValue('Y_SOUSPLAN4', vTobAna.GetValue('Y_SOUSPLAN4'));
    vTobOD.PutValue('Y_SOUSPLAN5', vTobAna.GetValue('Y_SOUSPLAN5'));
  end;

  {26.06.07 FQ18513 YMO}
  vTobOD.PutValue('Y_DATEREFEXTERNE', vTobAna.GetValue('Y_DATEREFEXTERNE'));

  //Mise en devise pivot
  ODDevise(vTobOD);

  //Insertion
  vTobOD.InsertDB(nil);

  //MAj Section
  MajSoldeSectionTOB(vTobAna, True);

  //Maj journaux
  MajJournalAnaTob(vTobOD, False, True);
end;

procedure TOF_CPREIMPUTANA.ODDevise(vTobOD: TOB);
begin
  if vTobOD.GetValue('Y_DEVISE') = V_PGI.DevisePivot then Exit;
  vTobOD.PutValue('Y_DEBITDEV', vTobOD.GetValue('Y_DEBIT'));
  vTobOD.PutValue('Y_CREDITDEV', vTobOD.GetValue('Y_CREDIT'));
  vTobOD.PutValue('Y_TAUXDEV', 1.0);
  vTobOD.PutValue('Y_DEVISE', V_PGI.DevisePivot);
end;

function TOF_CPREIMPUTANA.LaDateOk: boolean;
var
  DD: TDateTime;
  Err: integer;
begin
  result := False;
  if not IsValidDate(DATEGENERE.Text) then
  begin
    HM.Execute(5, '', '');
    DATEGENERE.Text := DateToStr(V_PGI.DateEntree);
    DateGene := V_PGI.DateEntree;
  end
  else
  begin
    DD := StrToDate(DATEGENERE.Text);
    Err := DateCorrecte(DD);
    if Err > 0 then
    begin
      HM.Execute(Err + 4, '', '');
      DATEGENERE.Text := DateToStr(V_PGI.DateEntree);
      DateGene := V_PGI.DateEntree;
    end
    else
    begin
      if RevisionActive(DD) then
      begin
        DATEGENERE.Text := DateToStr(V_PGI.DateEntree);
        DateGene := V_PGI.DateEntree;
      end
      else
      begin
        DateGene := DD;
        Result := True;
      end;
    end;
  end;
end;

procedure TOF_CPREIMPUTANA.DATEGENEREExit(Sender: TOBject);
begin
  if not LaDateOk then if DATEGENERE.CanFocus then DATEGENERE.SetFocus;
end;

function TOF_CPREIMPUTANA.JournalOk : boolean ;
var
  Q : TQuery ;
begin
  result:=True ;
  Q := OpenSQL('SELECT J_FERME, J_COMPTEURNORMAL, J_AXE from JOURNAL Where J_JOURNAL="' + JALGENERE.Value + '"', True);
  if not Q.EOF then
  begin
    if Q.Fields[0].AsString = 'X' then
    begin
      HM.Execute(6, '', '');
      result := false;
    end
    else if Q.Fields[1].AsString = '' then
    begin
      HM.Execute(7, '', '');
      result := false;
    end
    else if Q.Fields[2].AsString <> Y_AXE.Value then
    begin
      HM.Execute(8, '', '');
      result := false;
    END;
  end
  else result := false;
  Ferme(Q);
end ;

procedure TOF_CPREIMPUTANA.TraiteQte(vTobModif : TOB ; vTobOrigine : TOB ; ligneEncours : integer ; vTobVentil : TOB ; var TotalQte1, TotalQte2, TotalTauxQte1, TotalTauxQte2 : Double );
var
  TauxQte1, TauxQte2, dRatio, Qte1, Qte2   : Double ;
begin
  if RIMPUT.ItemIndex = 0 then dRatio := 1.0
  else dRatio := RATIO.Value;

  TauxQte1 := vTobVentil.Detail[ligneEncours].GetValue('V_TAUXQTE1');
  TauxQte2 := vTobVentil.Detail[ligneEncours].GetValue('V_TAUXQTE2');

  // Renseignement généraux
  vTobModif.PutValue('Y_QAULIFECRQTE1', vTobOrigine.GetValue('Y_QUALIFECRQTE1'));
  vTobModif.PutValue('Y_QUALIFECRQTE2', vTobOrigine.GetValue('Y_QUALIFECRQTE2'));
  vTobModif.PutValue('Y_QUALIFQTE1', vTobOrigine.GetValue('Y_QUALIFQTE1'));
  vTobModif.PutValue('Y_QUALIFQTE2', vTobOrigine.GetValue('Y_QUALIFQTE2'));

  vTobModif.PutValue('Y_POURCENTQTE1', Arrondi(TauxQte1, 4));
  vTobModif.PutValue('Y_POURCENTQTE2', Arrondi(TauxQte2, 4));

  // Totaux des qtes, reprise de la ligne géné ou à 0 ???? a voir...
  vTobModif.PutValue('Y_TOTALQTE1', Arrondi( vTobOrigine.GetValue('Y_QTE1') * dRatio, 3));
  vTobModif.PutValue('Y_TOTALQTE2', Arrondi( vTobOrigine.GetValue('Y_QTE2') * dRatio, 3));

  // Cumul des taux pour gestion des arrondi sur du 100%
  TotalTauxQte1 := TotalTauxQte1 + TauxQte1 ;
  TotalTauxQte2 := TotalTauxQte2 + TauxQte2 ;

  // GESTION QTE1
  if (vTobModif.GetValue('Y_TOTALQTE1') = 0.0) or (TauxQte1 = 0.0) then
    vTobModif.PutValue('Y_QTE1', 0.0)
  else
    begin
    // Quantités à répartir :
    Qte1 := vTobModif.GetValue('Y_TOTALQTE1');
    if ( ligneEncours = vTobVentil.Detail.Count - 1) and ( Arrondi( TotalTauxQte1 - 100.0 , 4 ) = 0 ) then
      // ligne N : gestion des arrondi si taux de 100% -> on calcul le reste a placer
      vTobModif.PutValue('Y_QTE1', Qte1 - TotalQte1)
    else
    // ligne 1 à n-1, application des taux sur qte à répartir
    begin
      Qte1 := Arrondi( Qte1 * (TauxQte1 / 100) , 3 );
      vTobModif.PutValue('Y_QTE1', Qte1);
      TotalQte1 := TotalQte1 + Qte1;
    end;
  end;

  // GESTION QTE2
  if (vTobModif.GetValue('Y_TOTALQTE2') = 0.0) or (TauxQte2 = 0.0) then
    vTobModif.PutValue('Y_QTE2', 0.0)
  else
    begin
    // Quantités à répartir :
    Qte2 := vTobModif.GetValue('Y_TOTALQTE2');
    if ( ligneEncours = vTobVentil.Detail.Count -1 ) and ( Arrondi( TotalTauxQte2 - 100.0 , 4 ) = 0 ) then
      // ligne N : gestion des arrondi -> on calcul le reste a placer
      vTobModif.PutValue('Y_QTE1',Qte2 - TotalQte2)
    else
      // ligne 1 à n-1, application des taux sur qte à répartir
      begin
      Qte2 := Arrondi( Qte2 * (TauxQte2 / 100) , 3 ) ;
      vTobModif.PutValue('Y_QTE2', Qte2);
      TotalQte2 := TotalQte2 + Qte2;
    end;
  end;

end;

procedure TOF_CPREIMPUTANA.BListePieceClick(Sender : TObject);
begin
  if TPieces.Count <= 0 then Exit;
  VisuPiecesGenere(TPieces,EcrAna,12);
end;


function TOF_CPREIMPUTANA.VerifModelRestriction(Q: TQuery): Boolean;
var
  i:           integer;
  Lst:         TStrings;
  RestrictAna: TRestrictionAnalytique;
  ParamAxe:    TParamAxe;
  sAxe:        string;
begin
  {b FP 29/12/2005}
  ParamAxe := RecherchePremDerAxeVentil;

  if ParamAxe.premier_axe = 0 then sAxe := Y_AXE.Value[2]
  else sAxe := IntToStr(ParamAxe.premier_axe);

  Result      := True;
  Lst         := TStringList.Create;
  RestrictAna := TRestrictionAnalytique.Create;

  if FListe.AllSelected then InitMove(Q.RecordCount, Ecran.Caption)
                        else InitMove(FListe.NbSelected, Ecran.Caption);
  try
    if FListe.AllSelected then
    begin
      Q.First;
      while not Q.eof do
      begin
        MoveCur(False);
        if Lst.IndexOf(Q.FindField('Y_GENERAL').AsString) = -1 then
          begin
          if not RestrictAna.IsCompteGeneAutorise(Q.FindField('Y_GENERAL').AsString,
                   'A'+sAxe, 'TY'+sAxe, VENTTYPE.Value) then
            begin
            Result := False;
            break;
            end
          else
            begin
            Lst.Add(Q.FindField('Y_GENERAL').AsString);
            end;
          end;
        Q.Next;
      end;
    end
    else
    begin
      for i := 0 to FListe.NbSelected - 1 do
      begin
        MoveCur(False);
        FListe.GotoLeBookmark(i);
        {b Thl 30/11/2006 }
        {$IFDEF EAGLCLIENT}
        Q.Seek(FListe.Row-1);
        {$ENDIF}
        {e Thl }
        if Lst.IndexOf(GetChamp(Q,'Y_GENERAL')) = -1 then
          begin
          if not RestrictAna.IsCompteGeneAutorise(GetChamp(Q,'Y_GENERAL'),
                   'A'+sAxe, 'TY'+sAxe, VENTTYPE.Value) then
            begin
            Result := False;
            break;
            end
          else
            begin
            Lst.Add(GetChamp(Q,'Y_GENERAL'));
            end;
          end;
      end;
    end;
  finally
    Lst.Free;
    RestrictAna.Free;
  end;
  {e FP 29/12/2005}
end;


{***********A.G.L.***********************************************
Auteur  ...... : Yann MORENO
Créé le ...... : 18/02/2008
Modifié le ... :   /  /
Description .. : Ajout du bouton 'tout sélectionner'
Mots clefs ... : FQ22310
*****************************************************************}
procedure TOF_CPREIMPUTANA.SlctAllClick(Sender: TObject); //YMO
var
  Fiche : TFMul;
begin
  Fiche := TFMul(Ecran);
  {$IFDEF EAGLCLIENT}
  if not Fiche.FListe.AllSelected then begin
    if not Fiche.FetchLesTous then Exit;
  end;
  {$ENDIF}
  Fiche.bSelectAllClick(nil);
end;

initialization
  registerclasses([TOF_CPREIMPUTANA]);
end.

