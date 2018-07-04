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
unit CLOTURESERVEUR;

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
    DBTables,
    UtilSoc,
{$ENDIF}
    hmsgbox,
    SaisUtil,
    SaisComm,
    HStatus,
    HSysMenu,
    Mask,
    CpteUtil,
    ParaClo,
    SoldeCpt,
    HTB97,
    HPanel,
    UiUtil,
    Paramsoc,
    ULIBCLOTURE, // Types et fonctions externes non graphiques
    uLanceProcess
    ;

{$IFDEF EAGLCLIENT}
    type tquery = TOB;
{$ENDIF}

//==================================================
// Externe
//==================================================
Function ClotureComptable(Definitive : Boolean) : Boolean;
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
        OnSort  : Boolean;
        Auto    : Boolean;
        OkAuto  : Boolean;


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
    end;


//================================================================================
// Implementation
//================================================================================
implementation

uses
{$IFDEF SPEC302}
    Societe,
{$ENDIF}
    UtilPgi
{$IFDEF EAGLCLIENT}
{$ELSE}
{$IFNDEF SANSCOMPTA}
  {$IFNDEF CCMP}
    ,ImoClo
  {$ENDIF}
{$ENDIF}
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
function ClotureComptable(Definitive : Boolean) : Boolean;
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

    FClo := TFClos.Create(Application);
    OutProg := false;
    FClo.CloDef := Definitive;
    FClo.OnSort := OutProg;
    FClo.Auto := false;
    FClo.Exo1 := VH^.EnCours;
    FClo.Exo2 := VH^.Suivant;

    if (Definitive) then FClo.Caption:=FClo.HMess.Mess[48] ;
    PP := FindInsidePanel;

    if ((PP=nil) or (true)) then
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
  ClotureProcess := TTraitementCloture.Create( self, Exo1.Code, Exo2.Code, CloDef, Auto ) ;

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
{$IFDEF SPEC302}
    FicheSociete(taModif,3,true);
{$ELSE}
  {$IFDEF EAGLCLIENT}
    // todo
    // non encore implementé en EAGL a cause de 'UtilSoc' !
  {$ELSE}
    ParamSociete(false,'','SCO_COMPTESSPECIAUX','',ChargeSocieteHalley,ChargePageSoc,SauvePageSocSansVerif,InterfaceSoc,1105000);
  {$ENDIF}
{$ENDIF}
    AlimCpt;
    RefO.SetFocus;
end;

procedure TFCloS.BParamClick(Sender: tobject);
begin
  ParametrageCloture( CloDef, ParaClo, ClotureProcess.GetExisteCptPV ) ;
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
{$IFDEF SPEC302}
Var Q: TQuery;
{$ENDIF}
begin
{$IFDEF SPEC302}
    Q := OpenSQL('SELECT SO_FERMEBIL, SO_OUVREBIL, SO_RESULTAT, SO_FERMEPERTE, SO_OUVREPERTE, SO_FERMEBEN, SO_OUVREBEN, SO_JALFERME, SO_JALOUVRE FROM SOCIETE WHERE SO_SOCIETE="'+V_PGI.CodeSociete+'"');
    Result := (not Q.Eof);
    if (not Q.Eof) then
      begin
      BilC.Text := Q.FindField('SO_FERMEBIL').AsString;
      ResC.Text := Q.FindField('SO_RESULTAT').AsString;
      PerC.Text := Q.FindField('SO_FERMEPERTE').AsString;
      BenC.Text := Q.FindField('SO_FERMEBEN').AsString;
      JalC.Text := Q.FindField('SO_JALFERME').AsString;
      BilO.Text := Q.FindField('SO_OUVREBIL').AsString;
      PerO.Text := Q.FindField('SO_OUVREPERTE').AsString;
      BenO.Text := Q.FindField('SO_OUVREBEN').AsString;
      JalO.Text := Q.FindField('SO_JALOUVRE').AsString;
      end;
    Ferme(Q) ;
{$ELSE}
    Result := true ;
    BilC.Text := GetParamSocSecur('SO_FERMEBIL','');
    ResC.Text := GetParamSocSecur('SO_RESULTAT','');
    PerC.Text := GetParamSocSecur('SO_FERMEPERTE','');
    BenC.Text := GetParamSocSecur('SO_FERMEBEN','');
    JalC.Text := GetParamSocSecur('SO_JALFERME','');
    BilO.Text := GetParamSocSecur('SO_OUVREBIL','');
    PerO.Text := GetParamSocSecur('SO_OUVREPERTE','');
    BenO.Text := GetParamSocSecur('SO_OUVREBEN','');
    JalO.Text := GetParamSocSecur('SO_JALOUVRE','');
{$ENDIF}
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
    ClotureOK : Boolean;
    BalOk     : Boolean;
    errID     : Integer ;
{$IFNDEF EAGLCLIENT}
  {$IFNDEF SANSCOMPTA}
    {$IFNDEF CCMP}
      ExoClo  : TExoDate;
    {$ENDIF}
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

    // -----------------------------------------------
    // ----- VERIFICATION DES FAUX COMPTES 6 ou 7 ----
    // -----------------------------------------------
    // Traitement
    errID := ClotureProcess.CompteChargeOk ;
    // Gestion des erreurs
    if (errID<>CLO_PASERREUR) then
      begin
      {$IFDEF CERIC}
        // BPY le 07/01/2004 => Demande CERIC
        if (HMess.Execute(64,'','') = mrNo) then exit;
      {$ELSE}
        // Attention : Certains comptes 6 ne sont pas de nature charge.
        HMess.Execute(errID,'','');
        // Vous devez modifier la nature de ces comptes.
        HMess.Execute(62,'','');
      {$ENDIF CERIC}
      end;
    // Traitement
    errID := ClotureProcess.CompteProduitOk ;
    // Gestion des erreurs
    if (errID<>CLO_PASERREUR) then
      begin
      {$IFDEF CERIC}
        // BPY le 07/01/2004 => Demande CERIC
        if (HMess.Execute(65,'','') = mrNo) then exit;
      {$ELSE}
        // Attention : Certains comptes 7 ne sont pas de nature produit.
        HMess.Execute(errID,'','');
        // Vous devez modifier la nature de ces comptes.
        HMess.Execute(62,'','');
      {$ENDIF CERIC}
      end;

{$IFDEF EAGLCLIENT}
{$ELSE}
{$IFNDEF SANSCOMPTA}
  {$IFNDEF CCMP}

    // --------------------------------------------------------
    // ----- VERIFICATION DE LA CLOTURE DES IMMOBILISATION ----
    // --------------------------------------------------------
    // Clôture des immobilisations dans l'option amortissement
    if (not Auto) then
      if (AuMoinsUneImmo and CloDef) then
        begin
        QuelDateDeExo(GetParamSocSecur('SO_EXOCLOIMMO',''),ExoClo);
        // si la date de dernière clôture immo est antérieure à l'encours, on lance la clôture des immos
        if VH^.Encours.Deb > ExoClo.Deb then
          begin
          PGIInfo ('Le traitement suivant va proposer la clôture impérative des immobilisations avant la clôture comptable.',Caption );
          AfficheClotureImmo;
          QuelDateDeExo(GetParamSocSecur('SO_EXOCLOIMMO',''),ExoClo);
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
      if ( not ParametrageCloture( CloDef, ParaClo, ClotureProcess.GetExisteCptPV ) ) then
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
    errID := ClotureProcess.Cloture ( ClotureOk, OkTraite ) ;
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

    // Gestion ecran
    EnableControls(Self,true);
    LPatience.Visible := false;
    ChangeEcran(true);

    // Message final
    if (not Auto) then
      if (ClotureOk) then
        HMess.Execute(CLO_CLOTUREOK,'',''); // Le traitement s'est correctement terminé.

end;

procedure TFCloS.ClotureProcessServer;
var i         : Integer;
    ClotureOK : Boolean;
    BalOk     : Boolean;
    errID     : Integer ;
    TobParam  : TOB ;
    TobResult : TOB ;
  {$IFNDEF SANSCOMPTA}
    {$IFNDEF CCMP}
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

    // -----------------------------------------------
    // ----- VERIFICATION DES FAUX COMPTES 6 ou 7 ----
    // -----------------------------------------------
    // Traitement
    errID := ClotureProcess.CompteChargeOk ;
    // Gestion des erreurs
    if (errID<>CLO_PASERREUR) then
      begin
      {$IFDEF CERIC}
        // BPY le 07/01/2004 => Demande CERIC
        if (HMess.Execute(65,'','') = mrNo) then exit;
      {$ELSE}
        // Attention : Certains comptes 6 ne sont pas de nature charge.
        HMess.Execute(errID,'','');
        // Vous devez modifier la nature de ces comptes.
        HMess.Execute(62,'','');
      {$ENDIF CERIC}
      end;
    // Traitement
    errID := ClotureProcess.CompteProduitOk ;
    // Gestion des erreurs
    if (errID<>CLO_PASERREUR) then
      begin
      {$IFDEF CERIC}
        // BPY le 07/01/2004 => Demande CERIC
        if (HMess.Execute(65,'','') = mrNo) then exit;
      {$ELSE}
        // Attention : Certains comptes 7 ne sont pas de nature produit.
        HMess.Execute(errID,'','');
        // Vous devez modifier la nature de ces comptes.
        HMess.Execute(62,'','');
      {$ENDIF CERIC}
      end;

  {$IFNDEF SANSCOMPTA}
    {$IFNDEF CCMP}
    // --------------------------------------------------------
    // ----- VERIFICATION DE LA CLOTURE DES IMMOBILISATION ----
    // --------------------------------------------------------
    if (not Auto) then
      if (AuMoinsUneImmo and CloDef) then
        begin
        QuelDateDeExo(GetParamSocSecur('SO_EXOCLOIMMO',True),ExoClo);
        // si la date de dernière clôture immo est antérieure à l'encours, on lance la clôture des immos
        if VH^.Encours.Deb > ExoClo.Deb then
          begin
          PGIInfo ('La clôture des immobilisations est impérative avant la clôture comptable.'
                   + #13#10 + ' Veuillez l''effectuer sur un poste 2 tiers', Caption );
          BValider.Visible := false;
          ModalResult := mrCancel;
          Exit ;
          end;
        end;
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
      if ( not ParametrageCloture( CloDef, ParaClo, ClotureProcess.GetExisteCptPV ) ) then
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

    // TobResult bien renseignée ?
    if TobResult.FieldExists('RESULT') then
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
    TobResult.Free ;
    TobParam.Free ;

    // -----------------------
    // ----- IT'S THE END ----
    // -----------------------

    // Gestion ecran
    AttenteServeur( false ) ;

    // Affichage incident
    if ErrId <> CLO_CLOTUREOK then
      if (not Auto) then
          HMess.Execute(errID,'','');

    // Message final
    if (not Auto) then
      if (ClotureOk) then
        HMess.Execute(CLO_CLOTUREOK,'',''); // Le traitement s'est correctement terminé.

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

end.

