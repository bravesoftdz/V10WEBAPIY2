{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 27/02/2003
Modifié le ... : 22/04/2003
Description .. : Fiche de cloture
Suite ........ : 22/04/2003 - CA - FQ 12159 - initilialisation correcte de
Suite ........ : E_ETATLETTRAGE sur compte non lettrable.
Suite ........ : 22/04/2003 - CA - FQ 12215 - mise à jour de DATECREATION
Suite ........ : 22/04/2003 - CA - FQ 12294 : mise à jour de l'exercice de référence
Suite ........ : 16/05/2003 - CA - FQ 12308 : Clôture sur un compte ventilable
Mots clefs ... : CLOTURE;FICHE
*****************************************************************}
unit CLOTURE;

//================================================================================
// Interface
//================================================================================
interface

uses
    Windows,
    Messages,
    SysUtils,
    Classes,
    Graphics,
    Controls,
    forms,
    Dialogs,
    StdCtrls,
    Buttons,
    ExtCtrls,
    Hcompte,
    Hctrls,
    ENT1,
    HEnt1,
    utoB,
{$IFDEF EAGLCLIENT}
{$ELSE}
    DB,
    {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
    UtilSoc, // FQ 17236
    hmsgbox,
    SaisUtil,
    SaisComm,
    ParaClo,
{$IFNDEF CCADM}
    CpteUtil,
    SoldeCpt,
{$ENDIF}
    ULIBCLOTURE, // Types et fonctions externes non graphiques
    HStatus,
    HSysMenu,
    Mask,
    HTB97,
    HPanel,
    UiUtil,
    Paramsoc,
    {$IFDEF NETEXPERT}
    UNEACtions,
    {$ENDIF}
    {$IFDEF MODENT1}
    CPTypeCons,
    {$IFDEF AMORTISSEMENT}
    CPProcMetier,
    {$ENDIF}
    {$ENDIF MODENT1}

    uLanceProcess, TntStdCtrls
    ;

{$IFDEF EAGLCLIENT}
    type tquery = TOB;
{$ENDIF}

//==================================================
// Externe
//==================================================
Function ClotureComptable(Definitive : Boolean ; vAnoDyna : boolean = false ) : Boolean;
Function SimuleCloture(Exo1,Exo2 : tExoDate ; CaptionF : String) : Boolean;

//==================================================
// Definition de class
//==================================================
type
    TFCloS = class(Tform)
        GBFerme: TGroupBox;
        TSO_FERMEBIL: THLabel;
        BilC: THCpteEdit;
        TSO_FERMEBEN: THLabel;
        TSO_FERMEPERTE: THLabel;
        TSO_JALFERME: THLabel;
        TSO_RESULTAT: THLabel;
        ResC: THCpteEdit;
        PerC: THCpteEdit;
        BenC: THCpteEdit;
        JalC: THCpteEdit;
        HLabel1: THLabel;
        RefC: TEdit;
        LibC: TEdit;
        HLabel2: THLabel;
        GBOuvre: TGroupBox;
        HLabel3: THLabel;
        HLabel4: THLabel;
        HLabel5: THLabel;
        HLabel6: THLabel;
        HLabel8: THLabel;
        HLabel9: THLabel;
        BilO: THCpteEdit;
        PerO: THCpteEdit;
        BenO: THCpteEdit;
        JalO: THCpteEdit;
        RefO: TEdit;
        LibO: TEdit;
        HPB: TtoolWindow97;
        BAide: TtoolbarButton97;
        BValider: TtoolbarButton97;
        BFerme: TtoolbarButton97;
        HMess: THMsgBox;
        GBTraitement: TGroupBox;
        GBencours: TGroupBox;
        EnCours: TLabel;
        GroupBox1: TGroupBox;
        HCpt1: TLabel;
        HCpt2: TLabel;
        GroupBox2: TGroupBox;
        Cpt1: TLabel;
        Cpt2: TLabel;
        Label1: TLabel;
        Label2: TLabel;
        Label3: TLabel;
        Label4: TLabel;
        Label5: TLabel;
        FAF1: TImage;
        FOK1: TImage;
        FAF2: TImage;
        FOK2: TImage;
        FAF3: TImage;
        FOK3: TImage;
        FAF4: TImage;
        FOK4: TImage;
        FAF5: TImage;
        FOK5: TImage;
        Panel1: TPanel;
        Label6: TLabel;
        XX: TLabel;
        Label8: TLabel;
        YY: TLabel;
        BSoc: TtoolbarButton97;
        BParam: TtoolbarButton97;
        HMTrad: THSystemMenu;
        HLabel10: THLabel;
        DateDebN1: TMaskEdit;
        Label7: TLabel;
        DateFinN1: TMaskEdit;
        HLabel12: THLabel;
        DateDebN: TMaskEdit;
        Label9: TLabel;
        DateFinN: TMaskEdit;
        PPatience: TPanel;
        H_TitreGuide: TLabel;
        PFenGuide: TPanel;
        Label10: TLabel;
        LPatience: TLabel;
        dock: Tdock97;
        Timer1: TTimer;

        procedure formShow(Sender: tobject);
        procedure formClose(Sender: tobject; var Action: TCloseAction);

        procedure BValiderClick(Sender: tobject);
        procedure BSocClick(Sender: tobject);
        procedure BParamClick(Sender: tobject);
        procedure BAideClick(Sender: tobject);
        procedure Timer1Timer(Sender: tobject);
        procedure ClotureProcessServer ;
        procedure Cloture2Tiers;
    procedure BFermeClick(Sender: TObject);
    private
        { Déclarations privées }

        // objet qui va faire tout le boulot ;)
        ClotureProcess : TTraitementCloture ;

        // Exo à cloturer et Exo suivant
        Exo1    : tExoDate;
        Exo2    : tExoDate;

        // Paramètres de cloture (cf paraclo.pas)
        ParaClo       : TParamCloture;

        // Indicateur de cloture définitive
        CloDef  : Boolean;

        // Pour traitement
        OnSort       : Boolean;
        Auto         : Boolean;
        OkAuto       : Boolean;
        AnoDyna      : boolean ;
        FClotureOK   : Boolean;

    // Fonctions utiles avec intéraction écran
        Function  VerifParamOk : boolean;

    // procédure graphique
        Function  AlimCpt : boolean;
        Procedure AlimParamCloture ;
        Procedure AttenteServeur ( debut : Boolean ) ;
    public
        { Déclarations publiques }
    // procédure graphique
        Procedure CursorSynchr ;
        Procedure Mess1(i : Integer);
        Procedure Mess2(i : Integer ; St1,St2 : String);
        procedure ChangeEcran(OkOk:Boolean);
        procedure InitGeneral ;
        procedure Patience ( vBoMode : Boolean ) ;
        procedure UpdateYY ( vStCaption : String ) ;
        procedure UpdateXX ( vStCaption : String ) ;
        procedure EtapeSuivante( vInEtape : Integer ) ;

        property  ClotureOk : boolean read FClotureOk ;

    end;


//================================================================================
// Implementation
//================================================================================
implementation

uses
  {$IFDEF MODENT1}
  ULibExercice,
  {$ENDIF MODENT1}
//  CPProcGen,       // AvertirMultiTable
  uLibRevision,
  UtilPgi

{$IFDEF AMORTISSEMENT}
    ,ImoClo
{$ENDIF}
    ;

{$R *.DFM}

//==================================================
// fonctions hors class : Point d'entré
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 27/02/2003
Modifié le ... : 05/03/2003
Description .. : Lancement d'une cloture definitive
Mots clefs ... : CLOTURE;DEFINITIVE
*****************************************************************}
function ClotureComptable(Definitive : Boolean ; vAnoDyna : boolean = false ) : Boolean;
var
    FClo: TFClos;
    OutProg : Boolean ;
    PP : THPanel ;
begin
    Result:=false ;

    if ((V_PGI.OutLook) and (VH^.Suivant.Code = '')) then
    begin
        HShowMessage('45;Clôture comptable;La clôture ne peut être faite car l''exercice suivant n''est pas ouvert.;E;O;O;O;','','') ;
        Exit ;
    end;

    if (not _BlocageMonoPoste(true)) then Exit ;

    FClo          := TFClos.Create(Application);
    OutProg       := false;
    FClo.CloDef   := Definitive;
    FClo.OnSort   := OutProg;
    FClo.Auto     := false;
    FClo.Exo1     := VH^.EnCours;
    FClo.Exo2     := VH^.Suivant;
    FClo.AnoDyna  := vAnoDyna ;

    if (Definitive) then FClo.Caption:=FClo.HMess.Mess[48] ;
    PP := FindInsidePanel;

    if ((PP=nil){ or (true)}) then
    begin
        try
            FClo.ShowModal;
        Finally
            OutProg := FClo.OnSort;
            FClo.free;
            _DeblocageMonoPoste(true);
        end;
    end
    else
    begin
        InitInside(FClo,PP);
        FClo.Show;
    end;

    if (OutProg and definitive) then Result := true;
    if FClo.AnoDyna then
     result := FClo.ClotureOk ;
    Screen.Cursor := SyncrDefault;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 27/02/2003
Modifié le ... : 05/03/2003
Description .. : Lancement d'une cloture simulé
Mots clefs ... : SIMULATION;CLOTURE
*****************************************************************}
function SimuleCloture(Exo1,Exo2 : tExoDate ; CaptionF : String) : Boolean;
var
    FClo: TFClos;
    PP : THPanel ;
begin
    Result := true;
    if (not _BlocageMonoPoste(true)) then Exit;

    FClo := TFClos.Create(Application);
    FClo.CloDef := true;
    FClo.OkAuto := true;
    FClo.Auto := true;
    FClo.Exo1 := Exo1;
    FClo.Exo2 := Exo2;
    FClo.Caption := CaptionF;
    PP := FindInsidePanel;

    if (PP = nil) then
    begin
        try
            FClo.ShowModal;
        finally
            Result := FClo.OkAuto;
            FClo.free;
            _DeblocageMonoPoste(true);
        end;
    end
    else
    begin
        InitInside(FClo,PP);
        FClo.Show;
    end;

    Screen.Cursor := SyncrDefault;
end;


//==================================================
// Fonction de la class : Evenement de la fiche
//==================================================
procedure TFCloS.formShow(Sender: tobject);
begin
  // contexte d'aide
  if Auto
    then HelpContext := 0
    else
      if CloDef
        then HelpContext := 7751000
        else HelpContext := 7742000;

  // Exo suivant non ouvert --> fermeture
  if (not Auto) then
    if (not V_PGI.OutLook) then
      if (Exo2.Code = '') then
        begin
        HMess.Execute(45,'','');
        PostMessage(Handle,WM_CLOSE,0,0);
        end;

  // MAJ des dates en fn des exo
  ExotoDates(Exo1.Code,DateDebN,DateFinN);
  ExotoDates(Exo2.Code,DateDebN1,DateFinN1);

  // MAJ écran
  AlimCpt;
  ChangeEcran(not Auto);

  // Instanciation processus
  ClotureProcess := TTraitementCloture.Create( self, Exo1.Code, Exo2.Code, CloDef, Auto , AnoDyna ) ;

  // Paramètre par défaut
  ParaClo.CloContrep      := EnPiedDC;
  ParaClo.ANOContrep      := EnPiedDC;
  ParaClo.CloPiece        := UneSeule;
  ParaClo.ANOPiece        := UneSeule;
  ParaClo.ANOCompteBanque := BQUneParRef;
  ParaClo.ANOComptePV     := SurAnalytique;

  // Init affichage des zones
  if (not CloDef) then
    begin
    GBFerme.Enabled := false;
    BilC.Enabled := false;
    JalC.Enabled := false;
    PerC.Enabled := false;
    BenC.Enabled := false;
    ResC.Enabled := false;
    LibC.Enabled := false;
    RefC.Enabled := false;
    HLabel1.Enabled := false;
    HLabel2.Enabled := false;
    TSO_RESULTAT.Enabled := false;
    TSO_FERMEBEN.Enabled := false;
    TSO_FERMEBIL.Enabled := false;
    TSO_FERMEPERTE.Enabled := false;
    TSO_JALFERME.Enabled := false;
    end;

  if (Auto) then
    begin
    Timer1.Enabled := true;
    BParam.Visible := false;
    BSoc.Visible := false;
    BValider.Visible := false;
    BFerme.Visible := false;
    BAide.Visible := false;
    end;
end;

procedure TFCloS.formClose(Sender: tobject; var Action: TCloseAction);
begin
  ClotureProcess.Free ;
  if (Parent is THPanel) then
  begin
    // GCO - 08/11/2007 - FQ 21804 - Mais pourquoi quelqu'un avait mis le
    // deblocage en commentaire ???
    _DeblocageMonoPoste(true);
{$IFDEF EAGLCLIENT}
{$ELSE}
    Action:=caFree ;
{$ENDIF}
  end;
end;

procedure TFCloS.BValiderClick(Sender: tobject);
begin
{$IFDEF EAGLCLIENT}
  ClotureProcessServer ;
{$ELSE}
  Cloture2Tiers ;
{$ENDIF}
end ;

procedure TFCloS.BAideClick(Sender: tobject);
begin
    CallHelptopic(Self);
end;

procedure TFCloS.BSocClick(Sender: tobject);
begin
  // FQ 17236 : mise en place pour le CWAS aussi
  ParamSociete(false,'','SCO_COMPTESSPECIAUX','',ChargeSocieteHalley,ChargePageSoc,SauvePageSocSansVerif,InterfaceSoc,1105000);
  AlimCpt;
  RefO.SetFocus;
end;

procedure TFCloS.BParamClick(Sender: tobject);
begin
  ParametrageCloture( CloDef, ParaClo, ClotureProcess.GetExisteCptPV ,false) ;
  // Affectation du résultat au process
  ClotureProcess.SetParamCloture ( ParaClo ) ;
end;

procedure TFCloS.Timer1Timer(Sender: tobject);
begin
    Timer1.Enabled := false;
    BValiderClick(nil);
    Close;
end;

//==================================================
// Fonction de la class : Gestion de messages
//==================================================
Procedure TFCloS.Mess1(i : Integer);
var lStTitre : String ;
begin
  if (Auto)
    then lStTitre := HMess.Mess[i+54]
    else lStTitre := HMess.Mess[i+6];
  // Gestion affichage cloture IFRS
  if ClotureProcess.EstModeIFRS
    then lStTitre := 'IFRS : ' + lStTitre ;
  // Modif du titre
  EnCours.Caption := lStTitre ;
  Application.ProcessMessages;
end;

Procedure TFCloS.Mess2(i : Integer ; St1,St2 : String);
begin
    Case i Of
        0 :
        begin
            HCpt1.Caption := St1;
            HCpt2.Caption := St2;
            Cpt1.Caption := '';
            Cpt2.Caption := '';
        end;
        1 :
        begin
            Cpt1.Caption := St1;
            Cpt2.Caption := St2;
        end;
    end;
    Application.ProcessMessages;
end;

//==================================================
// Fonction de la class : Autres fonctions
//==================================================

Function TFCloS.VerifParamOk : boolean;
var
    i,ii : Integer;
    St,FactClo,FactAN : String;
begin
    Result:=false ;

    i := VerifCpt(BilO.Text);
    if (i > 0) then
    begin
        St := ' '+Hmess.Mess[i];
        if (not Auto) then HMess.execute(16,'',st);
        Exit;
    end;

    i := VerifCpt(BenO.Text);
    if (i > 0) then
    begin
        St:=' '+Hmess.Mess[i];
        if (not Auto) then HMess.execute(17,'',st);
        Exit;
    end;

    i := VerifCpt(PerO.Text);
    if (i > 0) then
    begin
        St:=' '+Hmess.Mess[i];
        if (not Auto) then HMess.execute(18,'',st);
        Exit;
    end;

    i := VerifJal(JalO.Text,'ANO',FactAN);
    if (i > 0) then
    begin
        St:=' '+Hmess.Mess[i];
        if (not Auto) then HMess.execute(19,'',st);
        Exit;
    end;

    if (CloDef) then
    begin
        i := VerifCpt(BilC.Text);
        if (i > 0) then
        begin
            St := ' '+Hmess.Mess[i];
            if (not Auto) then HMess.execute(20,'',st);
            Exit;
        end;

        i := VerifCpt(BenC.Text);
        if (i > 0) then
        begin
            St := ' '+Hmess.Mess[i];
            if (not Auto) then HMess.execute(21,'',st);
            Exit;
        end;

        i := VerifCpt(PerC.Text);
        if (i > 0) then
        begin
            St := ' '+Hmess.Mess[i];
            if (not Auto) then if (not Auto) then HMess.execute(22,'',st);
            Exit;
        end;

        i := VerifJal(JalC.Text,'CLO',FactClo);
        if (i > 0) then
        begin
            St := ' '+Hmess.Mess[i];
            if (not Auto) then HMess.execute(23,'',st);
            Exit;
        end;

        i := VerifCpt(ResC.Text);
        if (i > 0) then
        begin
            St := ' '+Hmess.Mess[i];
            if (not Auto) then HMess.execute(24,'',st);
            Exit;
        end;
    end;

    Result:=true ;

    if (not Auto) then if (Multifacturier('ANO',factAN)) then
    begin
        ii := HMess.Execute(33,'','');
        Screen.Cursor := SyncrDefault;
        Case ii of
            mrNo,mrCancel :
            begin
                Result := false;
                Exit;
            end;
        end;
    end;

    if (not Auto) then if (CloDef) then if (Multifacturier('CLO',factClo)) then
    begin
        ii := HMess.Execute(32,'','');
        Screen.Cursor := SyncrDefault;
        Case (ii) of
            mrNo,mrCancel :
            begin
                Result := false;
                Exit;
            end;
        end;
    end;
end;

Function TFCloS.AlimCpt : boolean;
begin
    Result := true ;
    BilC.Text := GetParamSocSecur('SO_FERMEBIL', '', True);
    ResC.Text := GetParamSocSecur('SO_RESULTAT', '', True);
    PerC.Text := GetParamSocSecur('SO_FERMEPERTE', '', True);
    BenC.Text := GetParamSocSecur('SO_FERMEBEN', '', True);
    JalC.Text := GetParamSocSecur('SO_JALFERME', '', True);
    BilO.Text := GetParamSocSecur('SO_OUVREBIL', '', True);
    PerO.Text := GetParamSocSecur('SO_OUVREPERTE', '', True);
    BenO.Text := GetParamSocSecur('SO_OUVREBEN', '', True);
    JalO.Text := GetParamSocSecur('SO_JALOUVRE', '', True);
   if (not Auto) then
    if (not result) then HMess.execute(1,'','');
end;

procedure TFCloS.ChangeEcran(OkOk:Boolean) ;
begin
    GBTraitement.Visible := (not OkOk);
    GBEnCours.Visible := (not OkOk);
    GBOuvre.Visible := OkOk;
    GBFerme.Visible := OkOk;
    XX.Caption := '';
    YY.Caption := '';

    if (OkOk) then
    begin
        GBTraitement.Align := AlNone;
        GBEnCours.Align := AlNone;
        GBOuvre.Align := AlClient;
        GBFerme.Align := Altop;
    end
    else
    begin
        Cpt1.caption := '';
        Cpt2.Caption := '';
        HCpt1.Caption := '';
        HCpt2.Caption := '';
        GBTraitement.Align := Altop;
        GBEnCours.Align := AlClient;
        GBOuvre.Align := AlNone;
        GBFerme.Align := AlNone;
    end;
end;

procedure TFCloS.InitGeneral ;
begin
    FOK1.VISIBLE := false;
    FOK2.VISIBLE := false;
    FOK3.VISIBLE := false;
    FOK4.VISIBLE := false;
    FOK5.VISIBLE := false;
    FAF1.Visible := true;
    FAF2.Visible := true;
    FAF3.Visible := true;
    FAF4.Visible := true;
    FAF5.Visible := true;
end;

procedure TFCloS.CursorSynchr ;
begin
  Screen.Cursor := SyncrDefault;
end;

procedure TFCloS.Patience ( vBoMode : Boolean );
begin
    // Affichage fenêtre de patience
    PPatience.Visible := vBoMode ;
    if vBoMode then
      Application.ProcessMessages;
end;

procedure TFCloS.UpdateXX(vStCaption: String);
begin
  XX.Caption := vStCaption ;
end;

procedure TFCloS.UpdateYY(vStCaption: String);
begin
  YY.Caption := vStCaption ;
end;

procedure TFCloS.EtapeSuivante(vInEtape: Integer);
begin
  Case vInEtape Of
    1 : begin
        FAF1.Visible := false;
        FOK1.Visible := true;
        end ;
    2 : begin
        FAF2.Visible := false;
        FOK2.Visible := true;
        end ;
    3 : begin
        FAF3.Visible := false;
        FOK3.Visible := true;
        end ;
    4 : begin
        FAF4.Visible := false;
        FOK4.Visible := true;
        end ;
    5 : begin
        FAF5.Visible := false;
        FOK5.Visible := true;
        end ;
    end ;
end;

procedure TFCloS.AlimParamCloture;
begin
  if not Assigned(ClotureProcess) then Exit ;

  ClotureProcess.SetParamOuv ( RefO.Text, LibO.Text, JalO.Text,
                               BilO.Text, PerO.Text, BenO.Text,
                               StrToDate( DateDebN1.Text ), StrToDate( DateFinN1.Text ) ) ;

  ClotureProcess.SetParamFerm ( RefC.Text, LibC.Text, JalC.Text,
                                BilC.Text, PerC.Text, BenC.Text, ResC.Text ,
                                StrToDate( DateDebN.Text ), StrToDate( DateFinN.Text ) ) ;

end;

procedure TFCloS.Cloture2Tiers;
var i         : Integer;
    OkTraite  : Boolean;
   // ClotureOK : Boolean;
    BalOk     : Boolean;
    errID     : Integer ;
    {$IFDEF NETEXPERT}
    IsNetEXpert : Boolean;
    {$ENDIF}
{$IFNDEF EAGLCLIENT}
  {$IFNDEF SANSCOMPTA}
      {$IFDEF AMORTISSEMENT}
      ExoClo  : TExoDate;
      {$ENDIF AMORTISSEMENT}
  {$ENDIF}
{$ENDIF}
begin

    // -----------------------------------------------------
    // ----- ALIMENTATION DS PARAMETRES DANS LE PROCESS ----
    // -----------------------------------------------------
    AlimParamCloture ;
    {$IFDEF NETEXPERT}
    if not InterrogeUserService (IsNetEXpert) then
    begin
        if IsNetExpert then
        begin
            PGIInfo('Il y a des utilisateurs connectés sur le dossier Business Line, cette fonction n''est plus disponible',Caption);
            BValider.Visible := false;
            ModalResult := mrCancel;
            Exit;
        end;
    end;
    {$ENDIF}

    // -------------------------------------
    // ----- VERIFICATION DE LA BALANCE ----
    // -------------------------------------
    BalOk := True ;
    // Affichage du panneau de patience
    Patience( True );
    // Traitement
    errID := ClotureProcess.BalanceOk ;
    // Gestion des erreurs
    if errID <> CLO_PASERREUR then
      begin
      BalOk := false;
      if (errID = CLO_ERRBALEXO1) or (errID = CLO_ERRBALEXO1IFRS) then
        begin
        HMess.Execute(errID,'','');    // La balance de l'exercice à clôturer n'est pas équilibrée.
        // Peut plus valider...
        BValider.Visible := false;
        ModalResult := mrCancel;
        end ;
      end ;

    // Requete finie, on cache le panneau de patience
    Patience( false );
    if not BalOk then Exit ;

    // ================================
    // === Vérification des mvts GC ===
    // === SBO DEV3228 FQ 16676     ===
    // === CA FQ 19376              ===
    // ================================
    if ((CloDef) and (not ClotureProcess.CanCloseExoGC)) then
    // if not ClotureProcess.CanCloseExoGC then
    begin
      PGIBox('La clôture comptable ne peut être effectuée. Des pièces commerciales (pièces vivantes ou données à envoyer en compta) sont présentes sur cet exercice.', Caption) ;
      Exit ;
    end;

    // Vérification des comptes de charges
    // GCO - 04/12/2006 - FQ 18174
    errID := ClotureProcess.CompteChargeOk ;
    if (errID <> CLO_PASERREUR) then
    begin
      if PgiAsk(ClotureProcess.LastErrorMsg + '#10#13' +
         'Voulez-vous abandonner le traitement?', 'Clôture comptable') = MrYes then Exit;
    end;

    // Vérification des comptes de charges
    // GCO - 04/12/2006 - FQ 18174
    errID := ClotureProcess.CompteProduitOk ;
    if (errID <> CLO_PASERREUR) then
    begin
      if PgiAsk(ClotureProcess.LastErrorMsg + '#10#13' +
         'Voulez-vous abandonner le traitement?', 'Clôture comptable') = MrYes then Exit;
    end;

    // GCO - 25/01/2006 - FQ 13049
    errID := ClotureProcess.JalExtraComptableOk;
    if (errID <> CLO_PASERREUR) then
    begin
      if PgiAsk(ClotureProcess.LastErrorMsg + '#10#13' +
                'Voulez-vous abandonner le traitement?', 'Clôture comptable') = MrYes then Exit;
    end;

    errID := ClotureProcess.GeneExtraComptableOk;
    if (errID <> CLO_PASERREUR) then
    begin
      if PgiAsk(ClotureProcess.LastErrorMsg + '#10#13' +
                'Voulez-vous abandonner le traitement?', 'Clôture comptable') = MrYes then Exit;
    end;
    // FIN GCO

    // GCO - 04/12/2006 - FQ 18842
    errID := ClotureProcess.CompteBilanOk;
    if (errID <> CLO_PASERREUR) then
    begin
      if PgiAsk(ClotureProcess.LastErrorMsg + '#10#13' +
                'Voulez-vous abandonner le traitement?', 'Clôture comptable') = MrYes then Exit;
    end;

    // GCO - 26/09/2006 - PLUS de message NORME BOI en Cloture Provisoire
    // YMO 11/08/2006 Normes NF/BOI verif absence d'écritures non validées
    if CloDef then
    begin
      errID := ClotureProcess.EcritValidOk ;

      if (errID=CLO_ECRNOVALIDE) and
         (PgiAsk('En référence au BOI 13 L-1-06 N° 12 du 24 janvier 2006 paragraphe 27, nous vous rappelons ' + #13#10 +
                 'que vous devez valider les écritures avant la clôture d''exercice. Certaines écritures ' + #13#10 +
                 'courantes ne sont pas validées. Voulez-vous continuer ?', Caption) = MrNo) then Exit;
    end;

{$IFNDEF EAGLCLIENT}
  {$IFNDEF SANSCOMPTA}
      {$IFDEF AMORTISSEMENT}
    // --------------------------------------------------------
    // ----- VERIFICATION DE LA CLOTURE DES IMMOBILISATION ----
    // --------------------------------------------------------
    // Clôture des immobilisations dans l'option amortissement
    if (not Auto) then
      if (AuMoinsUneImmo and CloDef and VH^.OkModImmo) then     // FQ 13616 : clôture immo appelée uniquement si immo sérialisé
        begin
        QuelDateDeExo(GetParamSoc('SO_EXOCLOIMMO'),ExoClo);
        // si la date de dernière clôture immo est antérieure à l'encours, on lance la clôture des immos
        if VH^.Encours.Deb > ExoClo.Deb then
          begin
          PGIInfo ('Le traitement suivant va proposer la clôture impérative des immobilisations avant la clôture comptable.',Caption );
          AfficheClotureImmo;
          QuelDateDeExo(GetParamSoc('SO_EXOCLOIMMO'),ExoClo);
          // si la clôture des immos n'a pas été faite, on sort
          if VH^.Encours.Deb > ExoClo.Deb then
//      if (HMess.Execute(52,'','') <> mrYes) then
            begin
            BValider.Visible := false;
            ModalResult := mrCancel;
            Exit ;
            end;
          end;
        end;
    // Fin clôture des immobilisations dans l'option amortissement
  {$ENDIF}
  {$ENDIF}
{$ENDIF}


    // Si problème dans les tests préalable --> on quitte
    Screen.Cursor := SyncrDefault;

    // on passe à la suite -> Affichage fenêtre de patience
    Patience( false ) ;

{$IFNDEF CCS3}
    // --------------------------------------------------
    // ----- PARAMETRAGE DE LA CLOTURE POUR S5 ET S7 ----
    // --------------------------------------------------
    if (not Auto) then
      begin
      if ( not ParametrageCloture( CloDef, ParaClo, ClotureProcess.GetExisteCptPV ,AnoDyna) ) then
        Exit ;
      // Affectation du résultat au process
      ClotureProcess.SetParamCloture ( ParaClo ) ;
      end ;

{$ENDIF}

    // --------------------------
    // ----- 1° CONFIRMATION ----
    // --------------------------
    if (not Auto) then
      begin
      i := HMess.Execute(34,'','') ;
      // Sinon --> on quitte
      if (i <> mrYes) then
        Exit ;
      end ;

    // --------------------------------------
    // ----- Vérification des paramètres ----
    // --------------------------------------
    Screen.Cursor := SynCrDefault;
     if (not VerifParamOk) then
      Exit ;

    // --------------------------
    // ----- 2° CONFIRMATION ----
    // --------------------------
    if (not Auto) then
      begin
      i := HMess.Execute(35,'','') ;
      // Sinon --> on quitte
      if (i <> mrYes) then
        Exit ;
      end ;

    // ---------------------
    // ----- INIT ECRAN ----
    // ---------------------
    Screen.Cursor := SynCrDefault;
    ChangeEcran(false);
    EnableControls(Self,false);
    BValider.Visible := false;
    LPatience.Visible := true;

    // --------------------------------
    // ----- TRAITEMENT DE CLOTURE ----
    // --------------------------------
    errID := ClotureProcess.Cloture ( FClotureOk, OkTraite ) ;
    if (errID <> CLO_PASERREUR) then
      if (not Auto) then
          HMess.Execute(errID,'',''); // ATTENTION : Un incident s'est prouit pendant le traitement.La clôture va être annulée.

    // --------------------------------
    // ----- VERIFICATION GENERALE ----
    // --------------------------------
    if (OkTraite and (not Auto)) then
      ClotureProcess.VerifCloture ( OnSort )
    else
      OnSort:=false;

    // -----------------------
    // ----- IT'S THE END ----
    // -----------------------
    _DeblocageMonoPoste(true);

    // Gestion ecran
    EnableControls(Self,true);
    LPatience.Visible := false;
    ChangeEcran(true);

    // Message final
    if (not Auto) then
      if (ClotureOk) then
      begin
{$IFDEF NETEXPERT}
       // Synchronisation avec BL
        if IsNetExpert then
          SynchroOutBL ('C', DateDebN1.Text, DateFinN1.Text);
{$ENDIF}
        // GCO - 11/10/2007 - FQ 21633
        if VH^.Revision.Plan <> '' then
          SynchroRICAvecExercice;

        HMess.Execute(CLO_CLOTUREOK,'',''); // Le traitement s'est correctement terminé.
      end;
end;

procedure TFCloS.ClotureProcessServer;
var i         : Integer;
    ClotureOK : Boolean;
    BalOk     : Boolean;
    errID     : Integer ;
    TobParam  : TOB ;
    TobResult : TOB ;
  {$IFNDEF SANSCOMPTA}
  {$IFDEF AMORTISSEMENT}
      ExoClo  : TExoDate;
  {$ENDIF}
  {$ENDIF}
begin

    // -----------------------------------------------------
    // ----- ALIMENTATION DS PARAMETRES DANS LE PROCESS ----
    // -----------------------------------------------------
    AlimParamCloture ;

    // -------------------------------------
    // ----- VERIFICATION DE LA BALANCE ----
    // -------------------------------------
    BalOk := True ;
    // Affichage du panneau de patience
    Patience( True );
    // Traitement
    errID := ClotureProcess.BalanceOk ;
    // Gestion des erreurs
    if errID <> CLO_PASERREUR then
      begin
      BalOk := false;
      if ( errID = CLO_ERRBALEXO1 ) or ( errID = CLO_ERRBALEXO1IFRS ) then
        begin
        HMess.Execute(errID,'','');    // La balance de l'exercice à clôturer n'est pas équilibrée.
        // Peut plus valider...
        BValider.Visible := false;
        ModalResult := mrCancel;
        end ;
      end ;

    // Requete finie, on cache le panneau de patience
    Patience( false );
    if not BalOk then Exit ;

    // ================================
    // === Vérification des mvts GC ===
    // === SBO DEV3228 FQ 16676     ===
    // === CA FQ 18997              ===
    // ================================
    if ((CloDef) and (not ClotureProcess.CanCloseExoGC)) then
    begin
      PGIBox('La clôture comptable ne peut être effectuée. Des pièces commerciales (pièces vivantes ou données à envoyer en compta) sont présentes sur cet exercice.', Caption) ;
      Exit ;
    end ;

    // Vérification des comptes de charges
    // GCO - 04/12/2006 - FQ 18174
    errID := ClotureProcess.CompteChargeOk ;
    if (errID <> CLO_PASERREUR) then
    begin
      if PgiAsk(ClotureProcess.LastErrorMsg + '#10#13' +
         'Voulez-vous abandonner le traitement?', 'Clôture comptable') = MrYes then Exit;
    end;

    // Vérification des comptes de charges
    // GCO - 04/12/2006 - FQ 18174
    errID := ClotureProcess.CompteProduitOk ;
    if (errID <> CLO_PASERREUR) then
    begin
      if PgiAsk(ClotureProcess.LastErrorMsg + '#10#13' +
         'Voulez-vous abandonner le traitement?', 'Clôture comptable') = MrYes then Exit;
    end;

    // GCO - 25/01/2006 - FQ 13049
    errID := ClotureProcess.JalExtraComptableOk;
    if (errID <> CLO_PASERREUR) then
    begin
      if PgiAsk(ClotureProcess.LastErrorMsg + '#10#13' +
         'Voulez-vous abandonner le traitement?', 'Clôture comptable') = MrYes then Exit;
    end;

    errID := ClotureProcess.GeneExtraComptableOk;
    if (errID <> CLO_PASERREUR) then
    begin
      if PgiAsk(ClotureProcess.LastErrorMsg + '#10#13' +
         'Voulez-vous abandonner le traitement?', 'Clôture comptable') = MrYes then Exit;
    end;
    // FIN GCO

    // GCO - 04/12/2006 - FQ 18842
    errID := ClotureProcess.CompteBilanOk;
    if (errID <> CLO_PASERREUR) then
    begin
      if PgiAsk(ClotureProcess.LastErrorMsg + '#10#13' +
         'Voulez-vous abandonner le traitement?', 'Clôture comptable') = MrYes then Exit;
    end;

    // YMO 11/08/2006 Normes NF/BOI verif absence d'écritures non validées
    if CloDef then
    begin
      errID := ClotureProcess.EcritValidOk ;
      if (errID=CLO_ECRNOVALIDE) and
       (PgiAsk('En référence au BOI 13 L-1-06 N° 12 du 24 janvier 2006 paragraphe 27, nous vous rappelons ' + #13#10 +
               'que vous devez valider les écritures avant la clôture d''exercice. Certaines écritures ' + #13#10 +
               'courantes ne sont pas validées. Voulez-vous continuer ?') = MrNo) then Exit;
    end;

  {$IFNDEF SANSCOMPTA}
    {$IFDEF AMORTISSEMENT}
    // --------------------------------------------------------
    // ----- VERIFICATION DE LA CLOTURE DES IMMOBILISATION ----
    // --------------------------------------------------------
    // Clôture des immobilisations dans l'option amortissement
    if (not Auto) then
      if (AuMoinsUneImmo and CloDef and VH^.OkModImmo) then     // FQ 13616 : clôture immo appelée uniquement si immo sérialisé
        begin
        QuelDateDeExo(GetParamSoc('SO_EXOCLOIMMO'),ExoClo);
        // si la date de dernière clôture immo est antérieure à l'encours, on lance la clôture des immos
        if VH^.Encours.Deb > ExoClo.Deb then
          begin
          PGIInfo ('Le traitement suivant va proposer la clôture impérative des immobilisations avant la clôture comptable.',Caption );
          AfficheClotureImmo;
          QuelDateDeExo(GetParamSoc('SO_EXOCLOIMMO'),ExoClo);
          // si la clôture des immos n'a pas été faite, on sort
          if VH^.Encours.Deb > ExoClo.Deb then
//      if (HMess.Execute(52,'','') <> mrYes) then
            begin
            BValider.Visible := false;
            ModalResult := mrCancel;
            Exit ;
            end;
          end;
        end;
    // Fin clôture des immobilisations dans l'option amortissement
      {$ENDIF}
  {$ENDIF}

    // Si problème dans les tests préalable --> on quitte
    Screen.Cursor := SyncrDefault;
    // on passe à la suite -> Affichage fenêtre de patience
    Patience( false ) ;

{$IFNDEF CCS3}
    // --------------------------------------------------
    // ----- PARAMETRAGE DE LA CLOTURE POUR S5 ET S7 ----
    // --------------------------------------------------
    if (not Auto) then
      begin
      if ( not ParametrageCloture( CloDef, ParaClo, ClotureProcess.GetExisteCptPV,false ) ) then
        Exit ;
      // Affectation du résultat au process
      ClotureProcess.SetParamCloture ( ParaClo ) ;
      end ;
{$ENDIF}

    // --------------------------
    // ----- 1° CONFIRMATION ----
    // --------------------------
    if (not Auto) then
      begin
      i := HMess.Execute(34,'','') ;
      // Sinon --> on quitte
      if (i <> mrYes) then
        Exit ;
      end ;

    // --------------------------------------
    // ----- Vérification des paramètres ----
    // --------------------------------------
    Screen.Cursor := SynCrDefault;
     if (not VerifParamOk) then
      Exit ;

    // --------------------------
    // ----- 2° CONFIRMATION ----
    // --------------------------
    if (not Auto) then
      begin
      i := HMess.Execute(35,'','') ;
      // Sinon --> on quitte
      if (i <> mrYes) then
        Exit ;
      end ;

    // ---------------------
    // ----- INIT ECRAN ----
    // ---------------------
    Screen.Cursor := SynCrDefault;
    // Affichage du panneau de patience
    AttenteServeur( True ) ;

    // Préparation de la tob
    TobParam := ClotureProcess.CreerTobParamCloture ;
    TobParam.AddChampSupValeur('ONSORT', OnSort) ;

    // Traitement
    TobResult := LanceProcessServer('cgiCloture', 'cloture', 'aucun', TobParam, True ) ;
//      TobResult := LanceProcessLocal('cgiCloture', 'cloture', 'aucun', TobParam, True ) ;

    // TobResult bien renseignée ?
    if Assigned(TobResult) and TobResult.FieldExists('RESULT') then
      begin
      // Récupération du résultat
      errID     := TobResult.GetValue('RESULT') ;
      OnSort    := TobResult.GetValue('ONSORT') ;
      ClotureOk := TobResult.GetValue('CLOTUREOK') ;
      end
    else
      begin
      // Pb avec le process server
      errID     := CLO_ERRPROCESSSERVER ;
      OnSort    := False ;
      ClotureOk := False ;
      end ;

    // Libération mémoire
    if Assigned(TobResult) then
      TobResult.Free ;
    if Assigned(TobParam) then
      TobParam.Free ;

    // -----------------------
    // ----- IT'S THE END ----
    // -----------------------
    _DeblocageMonoPoste(true);

    // Gestion ecran
    AttenteServeur( false ) ;

    // Affichage incident
    if ErrId <> CLO_CLOTUREOK then
      if (not Auto) then
          HMess.Execute(errID,'','');

  // --------------------------------------------------------------------------------------
  // ---- On averti le serveur de la mise à jour des paramSoc et de la table exercice  ----
  // --------------------------------------------------------------------------------------
  AvertirCacheServer( 'PARAMSOC' ) ; // SBO : 30/06/2004
  AvertirCacheServer( 'EXERCICE' ) ; // SBO : 30/06/2004

  AvertirMultiTable('TTEXERCICE');   // GCO - 02/01/2008 - FQ 22048

  // GCO - 02/08/2007 - FQ 20648
  // AvertirCacherServer ne sert à rien, si derrière on ne recharge pas VH^.
  ChargeMagExo( False );

  // GCO - 11/10/2007 - FQ 21633
  if VH^.Revision.Plan <> '' then
    SynchroRICAvecExercice;

  // Message final
  if (not Auto) and (ClotureOk) then
  begin
    HMess.Execute(CLO_CLOTUREOK,'',''); // Le traitement s'est correctement terminé.
  end;  

end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 19/11/2003
Modifié le ... :   /  /
Description .. : Affiche / Cache une fenêtre modale d'attente pendant le
Suite ........ : traitement du process Serveur.
Suite ........ : Le paramètre debut indique si le message doît apparaître
Suite ........ : ou disparâitre
Mots clefs ... :
*****************************************************************}
procedure TFCloS.AttenteServeur(debut: Boolean);
begin

  EnableControls( Self, not debut );

  H_TitreGuide.Caption := 'Exécution de la cloture en cours' ;
  Label10.Caption      := 'Veuillez patienter...' ;

  Patience( debut ) ;

end;

procedure TFCloS.BFermeClick(Sender: TObject);
begin
  //SG6 06/12/2004 FQ 14987
  Close;
  if IsInside(Self) then
    CloseInsidePanel(Self);
end;

end.

