unit AssistRazActivite;

interface

uses Windows,
     Messages,
     SysUtils,                          
     Classes,
{$IFDEF EAGLCLIENT}
     MaineAGL,
{$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     FE_Main,
     HDB,
{$ENDIF}
     Graphics,
     Controls,
     Forms,
     Dialogs,
     assist,
     HSysMenu,
     hmsgbox,
     StdCtrls,
     ComCtrls,
     ExtCtrls,
     HTB97,
     Hctrls,
     HPanel,
     Mask,
     DBCtrls,
     HFLabel,
     UtilUtilitaires, TntStdCtrls, TntComCtrls, TntExtCtrls;

    procedure Assist_RazActivite;

type
  TFAssistRazActivite = class(TFAssist)
    TSMotDePasse: TTabSheet;
    TSConfirmation: TTabSheet;
    TSRapport: TTabSheet;
    RapportExec: TListBox;
    TSOptions: TTabSheet;
    GBMotDePasse: TGroupBox;
    Lib1: THLabel;
    MotDePasse: THCritMaskEdit;
    TMotDePasse: THLabel;
    GBOptions: TGroupBox;
    CBSuppDepots: TCheckBox;
    GBConfirmation: TGroupBox;
    Lib4: TFlashingLabel;
    Lib5: THLabel;
    Lib7: THLabel;
    Lib2: THLabel;
    Lib3: THLabel;
    Lib6: THLabel;
    bImprimer: TToolbarButton97;
    //Evènements de la forme
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure MdpOnChange(Sender: TObject);
    //Gestion des boutons
    procedure bSuivantClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    function TestPassword : boolean;
    procedure bImprimerClick(Sender: TObject);

  private
    //Tests
    function ConfAvantExecute(Titre, Msg : integer; Chaine : string) : boolean;
    //Exécution
    procedure ExecutionRaz;
    procedure MajRaz;

  public

  end;

const
TexteMessage: array[1..8] of string =(
          {1}  'ERREUR',
          {2}  'ATTENTION !!!',
          {3}  'CONFIRMATION',
          {4}  'Le mot de passe est incorrect.',
          {5}  'Il faut impérativement une sauvegarde récente de votre base '+
               'avant de continuer.'+#13+#10+'Si vous n''en avez pas, cliquez sur '+
               '"Abandonner" pour annuler l''opération sinon, cliquez sur "Oui" pour continuer.',
          {6}  'Veuillez confirmer l''éxécution de la remise à zéro de l''activité',
          {7}  'Cet utilitaire supprime définitivement toutes les informations '+
               'liées à la société.'+#13+#10+'Veuillez confirmer...',
          {8}  'L''éxécution de cet utilitaire est impossible tant qu''il y a '+
               'd''autres utilisateurs connectés sur la base.'
          );

implementation

uses HEnt1,
     UTOB,
     HStatus,
     ParamSoc,
     Ent1,
     LicUtil;

var EventsLog : TStringList;
    TobDesTables : TOB;
    MsgErr,DepotDefaut : string;

{$R *.DFM}

procedure Assist_RazActivite;
var RazAct_Assist : TFAssistRazActivite;
Begin
    if not BlocageMonoPoste(True) then exit;
    RazAct_Assist := TFAssistRazActivite.Create(Application);
    try
        RazAct_Assist.ShowModal;
    finally
        RazAct_Assist.free;
    end;
end;

{===============================================================================
 ============================= Evenements de la forme ==========================
 ===============================================================================}
procedure TFAssistRazActivite.FormShow(Sender: TObject);
begin
    inherited;
    bPrecedent.Enabled := False;
    bSuivant.Enabled := False;
    bFin.Enabled := False;
    bImprimer.Enabled := False;
    DepotDefaut := GetParamSoc('SO_GCDEPOTDEFAUT');
    if DepotDefaut = '' then
        CBSuppDepots.Enabled := False
        else
        CBSuppDepots.Enabled := True;
    TobDesTables := TOB.Create('DETABLES', nil, -1);
    EventsLog := TStringList.Create;
end;

procedure TFAssistRazActivite.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    DeblocageMonoPoste(True);
    TobDesTables.Free;
    EventsLog.Free;
    RapportExec.Free;
    P.Free;
end;

procedure TFAssistRazActivite.MdpOnChange(Sender: TObject);
begin
    inherited;
    if TestPassword then
        bSuivant.Enabled := True
        else
        bSuivant.Enabled := False;
end;

function TFAssistRazActivite.TestPassword : boolean;
begin
    if UpperCase(MotDePasse.Text) <> DayPass(Date) then
        Result := False
        else
        Result := True;
end;

{===============================================================================
 ============================= Gestion des boutons =============================
 ===============================================================================}
procedure TFAssistRazActivite.bSuivantClick(Sender: TObject);
begin
    inherited;
    if P.ActivePage.Name = 'TSOptions' then
    begin
        bSuivant.Enabled := True;
        bPrecedent.Enabled := True;
        bFin.Enabled := False;
    end else
    if P.ActivePage.Name = 'TSConfirmation' then
    begin
        if CBSuppDepots.Checked then
            Lib6.Caption := TraduireMemoire('Vous avez choisi du supprimer les dépôts.')
            else
            Lib6.Caption := '';
        bSuivant.Enabled := False;
        bPrecedent.Enabled := True;
        bFin.Enabled := True;
    end;
end;

procedure TFAssistRazActivite.bPrecedentClick(Sender: TObject);
begin
    inherited;
    if P.ActivePage.Name = 'TSConfirmation' then
    begin
        bSuivant.Enabled := False;
        bPrecedent.Enabled := True;
        bFin.Enabled := False;
    end else
    if P.ActivePage.Name = 'TSOptions' then
    begin
        bSuivant.Enabled := True;
        bPrecedent.Enabled := True;
        bFin.Enabled := False;
    end;
end;

procedure TFAssistRazActivite.bFinClick(Sender: TObject);
begin
    inherited;
    if bFin.Caption = 'Fin' then
    begin
        if not ConfAvantExecute(2,5,';W;YA;A;A;') then exit;
        if not ConfAvantExecute(3,6,';W;YN;N;N;') then exit;
        if not ConfAvantExecute(3,7,';W;YN;N;N;') then exit;
        ExecutionRaz;
    end else
    begin
        Close;
    end;
end;

procedure TFAssistRazActivite.bImprimerClick(Sender: TObject);
var TobToPrint : TOB;
    Cpt : integer;
begin
  inherited;
  TobToPrint := TOB.Create('',nil,-1);
  for Cpt := 0 to RapportExec.Items.Count -1 do
    UtilTobCreat(TobToPrint,'','',RapportExec.Items[Cpt],'');
  UtilTobPrint(TobToPrint,Caption,0);
  TobToPrint.free;
end;
{===============================================================================
 ============================= Confirmations ===================================
 ===============================================================================}
function TFAssistRazActivite.ConfAvantExecute(Titre, Msg : integer; Chaine : string) : boolean;
begin
    Result := False;
    if HShowMessage('0;'+TraduireMemoire(TexteMessage[Titre])+';'+
                         TraduireMemoire(TexteMessage[Msg])+Chaine,'','') = mrYes then
        Result := True;
end;

{===============================================================================
 ============================= Exécution =======================================
 ===============================================================================}
procedure TFAssistRazActivite.ExecutionRaz;
var TobJnal : TOB;
    NumEvt : integer;
    Qry : TQuery;
    io : TIOErr;
    Msg : string;
begin
    P.SelectNextPage(True);
    bSuivant.Enabled := False;
    bPrecedent.Enabled := False;
    bAnnuler.Enabled := False;
    bFin.Caption := 'Quitter';
    InitMove(20,'');
    MoveCur(False);
    TobDesTables.LoadDetailFromSQL('SELECT * FROM DETABLES');
    TobJnal := TOB.Create('JNALEVENT', nil, -1);
    io := Transactions(MajRaz,0);
    case io of
             oeOk : begin
                        TobJnal.PutValue('GEV_ETATEVENT', 'OK');
                        Msg := TraduireMemoire('Le traitement s''est effectué correctement.');
                        EventsLog.Add(Msg);
                        RapportExec.Items.Add(Msg);
                    end;
        oeUnknown : begin
                        TobJnal.PutValue('GEV_ETATEVENT', 'ERR');
                        EventsLog.Add(MsgErr);
                        RapportExec.Items.Add(MsgErr);
                        Msg := TraduireMemoire('*** LE TRAITEMENT A ETE ANNULE ***');
                        EventsLog.Add(Msg);
                        RapportExec.Items.Add(Msg);
                    end;
    end;

    //Maj journal d'évènements
    NumEvt := 0;
    Qry := OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT', True,-1,'',true);
    if not Qry.Eof then
        NumEvt := Qry.Fields[0].AsInteger;
    Ferme(Qry);
    inc(NumEvt);
    TobJnal.PutValue('GEV_NUMEVENT', NumEvt);
    TobJnal.PutValue('GEV_TYPEEVENT', 'UTI');
    TobJnal.PutValue('GEV_LIBELLE', TraduireMemoire('Remise à zéro de l''activité'));
    TobJnal.PutValue('GEV_DATEEVENT', Date);
    TobJnal.PutValue('GEV_UTILISATEUR', V_PGI.User);
    TobJnal.PutValue('GEV_BLOCNOTE',EventsLog.Text);
    TobJnal.InsertDB(nil);
    TobJnal.Free;
    FiniMove;
    bImprimer.Enabled := True;
end;

procedure TFAssistRazActivite.MajRaz;
var Qte : integer;
    Msg,LibTable,Sql : string;
    TobTmp : TOB;

    procedure DeleteLaTable(Table, Condition : string);
    begin
        MoveCur(False);
        TobTmp := TobDesTables.FindFirst(['DT_NOMTABLE'],[Table],True);
        if TobTmp <> nil then
        begin
            LibTable := TobTmp.GetValue('DT_LIBELLE');
            MSgErr := TraduireMemoire('*** ERREUR LORS DE LA SUPPRESSION DE "')+LibTable+'" ***';
            Qte := ExecuteSql('DELETE FROM '+Table+Condition);
            if Qte >= 0 then
            begin
                Msg := TraduireMemoire('Suppression de ')+IntToStr(Qte)+' '+
                       TraduireMemoire('lignes(s) dans "')+LibTable+'".';
                EventsLog.Add(Msg);
                RapportExec.Items.Add(Msg);
            end;
        end;
    end;

begin
    DeleteLaTable('ACOMPTES', '');
    DeleteLaTable('ADRESSES',' WHERE ADR_TYPEADRESSE = "PIE"');
    DeleteLaTable('COMPTADIFFEREE', '');
    DeleteLaTable('CTRLCAISMT', '');
    DeleteLaTable('DISPO', '');
    DeleteLaTable('DISPOLOT', '');
    DeleteLaTable('DISPOCONTREM', '');
    DeleteLaTable('DISPOSERIE', '');
    DeleteLaTable('JOURSCAISSE', '');
    DeleteLaTable('JOURSETAB', '');
    DeleteLaTable('JOURSETABEVT', '');
    DeleteLaTable('LIENSOLE', ' WHERE LO_TABLEBLOB <> "GA"');
    DeleteLaTable('LIGNE', '');
    DeleteLaTable('LIGNECOMPL', '');
    DeleteLaTable('LIGNELOT', '');
    DeleteLaTable('LIGNENOMEN', '');
    DeleteLaTable('LIGNESERIE', '');
    DeleteLaTable('LISTEINVENT', '');
    DeleteLaTable('LISTEINVLIG', '');
    DeleteLaTable('LISTEINVLIGCONTREM', '');
    DeleteLaTable('LISTEINVLOT', '');
    DeleteLaTable('METHODEAPPRO', ' WHERE GMA_FORMULEQTEVTE = "" AND GMA_FORMULEQTEACH = ""');
    DeleteLaTable('PIECE', '');
    DeleteLaTable('PIECEADRESSE', '');
    DeleteLaTable('PIEDBASE', '');
{$IFDEF BTP}
    DeleteLaTable('AFFAIRE', '');
    DeleteLaTable('AFFTIERS', '');
    DeleteLaTable('BDETETUDE', '');
    DeleteLaTable('BLIENEXCEL', '');
    DeleteLaTable('BMEMORISATION', '');
    DeleteLaTable('BTMEMOFACTURE', '');
    DeleteLaTable('BFAMILLECHANGE', '');
    DeleteLaTable('BSITUATIONS', '');
    DeleteLaTable('BSTRDOC', '');
    DeleteLaTable('BTPARDOC', ' WHERE BPD_NUMPIECE > 0');
    DeleteLaTable('BTPIECEMILIEME', '');
    DeleteLaTable('BVARDOC', '');
    DeleteLaTable('BVARIABLES', '');
    DeleteLaTable('CONSOMMATIONS', '');
    DeleteLaTable('DECISIONACH', '');
    DeleteLaTable('DECISIONACHLIG', '');
    DeleteLaTable('LIGNEBASE', '');
    DeleteLaTable('LIENDEVCHA', '');
    DeleteLaTable('LIGNEOUV', '');
    DeleteLaTable('LIGNEOUVPLAT', '');
    DeleteLaTable('LIGNEPHASES', '');
    DeleteLaTable('LIGNEFAC', '');
    DeleteLaTable('PHASESCHANTIER', '');
    DeleteLaTable('PIECERG', '');
    DeleteLaTable('PIEDBASERG', '');

    //FV1 : 05/02/2014 - FS#858 - En raz activité, la table PIECETRAIT n'est pas traitée
    DeleteLaTable('PIECETRAIT', '');
    //

    //FV1 - 03092012 => FS#71 - réinitialisation des tables de la GRC
    DeleteLaTable('RTDOCUMENT', '');
    DeleteLaTable('RTETATPLANNING', '');
    DeleteLaTable('RTPARAMPLANNING', '');
    DeleteLaTable('RTINFOS001', '');
    DeleteLaTable('RTINFOS002', '');
    DeleteLaTable('RTINFOS005', '');
    DeleteLaTable('RTINFOS007', '');
    DeleteLaTable('RTINFOS008', '');
    DeleteLaTable('RTINFOS009', '');
    //
    DeleteLaTable('RTINFOS00B', '');
    DeleteLaTable('RTINFOS00D', '');
    DeleteLaTable('RTINFOS00H', '');
    DeleteLaTable('RTINFOS00J', '');
    DeleteLaTable('RTINFOS00K', '');
    DeleteLaTable('RTINFOS00N', '');
    DeleteLaTable('RTINFOS00O', '');
    DeleteLaTable('RTINFOS00Q', '');
    DeleteLaTable('RTINFOS00V', '');
    DeleteLaTable('RTINFOS00W', '');
    DeleteLaTable('RTINFOS00X', '');
    DeleteLaTable('RTINFOS00Y', '');
    DeleteLaTable('RTINFOS00Z', '');
    DeleteLaTable('RTINFOSDATA', '');
    DeleteLaTable('RTINFOSDESC', '');
    //
    DeleteLaTable('ACTIONINTERVENANT', '');
    DeleteLaTable('PERSPHISTO', '');
    //DeleteLaTable('PROSPECTS', '');
    DeleteLaTable('ACTIONPIECE', '');
    DeleteLaTable('ACTIONS', '');
    DeleteLaTable('ACTIONSCHAINEES', '');
    DeleteLaTable('ACTIONSGENERIQUES', '');
    DeleteLaTable('CHAINAGEPIECES', '');
    DeleteLaTable('CIBLAGEELEMENT', '');
    DeleteLaTable('CIBLES', '');
    DeleteLaTable('GROUPECHAINAGE', '');
    DeleteLaTable('LIENBCONNAISSANCE', '');
    DeleteLaTable('OPERATIONS', '');
    DeleteLaTable('PARACTIONS', '');
    DeleteLaTable('PARCHAINAGES', '');
    DeleteLaTable('PERSPECTIVES', '');
    DeleteLaTable('PERSPECTIVESTIERS', '');
    DeleteLaTable('PERSHISTO', '');
    DeleteLaTable('PROJETS', '');
    DeleteLaTable('RQDEMDEROG', '');
    DeleteLaTable('RQPLANCORR', '');
    DeleteLaTable('RQTCRITERES', '');

{$ENDIF}
    DeleteLaTable('PIEDECHE', '');
    DeleteLaTable('PIEDPORT', '');
    DeleteLaTable('REMISELIGNE', '');
    DeleteLaTable('TRANSINVENT', '');
    DeleteLaTable('TRANSINVLIG', '');
    DeleteLaTable('VENTANA', ' WHERE YVA_TABLEANA="GL" OR YVA_TABLEANA="GP"');

    //Traitement de METHODEAPPRO
    TobTmp := TobDesTables.FindFirst(['DT_NOMTABLE'],['METHODEAPPRO'],True);
    if TobTmp <> nil then
    begin
        LibTable := TobTmp.GetValue('DT_LIBELLE');
        MsgErr := TraduireMemoire('*** ERREUR LORS DE LA MISE A JOUR DE "')+LibTable+'" ***';
        Sql := 'UPDATE METHODEAPPRO SET '+
               'GMA_DECLENCHEAPPRO = "", '+
               'GMA_QUANTITEAPPRO = 0, '+
               'GMA_DELAIOBTENTION = 0, '+
               'GMA_CLASSEABC = "", '+
               'GMA_CLASSEABCOLD = "", '+
               'GMA_PERIODICITE = "", '+
               'GMA_NBPERIODE = 0, '+
               'GMA_DATEAPPRO = "01/01/1900", '+
               'GMA_QTEECO = 0';
        Qte := ExecuteSQL(Sql);
        if Qte >= 0 then
            begin
                Msg := TraduireMemoire('Mise à jour de ')+IntToStr(Qte)+' '+
                       TraduireMemoire('enregistrement(s) dans "')+LibTable+'".';
                EventsLog.Add(Msg);
                RapportExec.Items.Add(Msg);
            end;
    end;

    //Traitement des dépôts et emplacements
    if CBSuppDepots.Checked then
    begin
        DeleteLaTable('DEPOTS', ' WHERE GDE_DEPOT <> "'+DepotDefaut+'"');
        DeleteLaTable('EMPLACEMENT', '');
        MsgErr := TraduireMemoire('*** ERREUR LORS DE LA MISE A JOUR DES EMPLACEMENTS DES ARTICLES ***');
        Qte := ExecuteSQL('UPDATE ARTICLE SET GA_TYPEEMPLACE = "" WHERE GA_TYPEEMPLACE <> ""');
        if Qte >= 0 then
            begin
                Msg := TraduireMemoire('Mise à jour de ')+IntToStr(Qte)+
                       TraduireMemoire(' article(s) effectuée.');
                EventsLog.Add(Msg);
                RapportExec.Items.Add(Msg);
            end;
    end else
    begin
        MsgErr := TraduireMemoire('*** ERREUR LORS DE LA MISE A JOUR DES EMPLACEMENTS ***');
        Qte := ExecuteSQL('UPDATE EMPLACEMENT SET GEM_EMPLACEOCCUPE = "-" WHERE GEM_EMPLACEOCCUPE = "X"');
        if Qte >= 0 then
            begin
                Msg := TraduireMemoire('Mise à jour de ')+IntToStr(Qte)+
                       TraduireMemoire(' emplacement(s) effectuée.');
                EventsLog.Add(Msg);
                RapportExec.Items.Add(Msg);
            end;
    end;

    //Maj date cloture stock
    MsgErr := TraduireMemoire('*** ERREUR LORS DE LA MISE A JOUR DE LA DATE DE CLOTURE DU STOCK ***');
    SetParamSoc('SO_GCDATECLOTURESTOCK', VH^.Encours.Deb);
    Msg := TraduireMemoire('Mise à jour de la date de clôture du stock effectuée.');
    EventsLog.Add(Msg);
    RapportExec.Items.Add(Msg);

    //Réinit des compteurs
    MsgErr := TraduireMemoire('*** ERREUR LORS DE LA REINITIALISATION DU COMPTEUR DES PIECES ***');
    Qte := ExecuteSQL('UPDATE SOUCHE SET SH_NUMDEPART = 1 WHERE SH_TYPE = "GES"');
    if Qte >= 0 then
        begin
            Msg := TraduireMemoire('Réinitialisation du compteur des pièces effectuée.');
            EventsLog.Add(Msg);
            RapportExec.Items.Add(Msg);
        end;

    ExecuteSQL ('UPDATE BCPTMENSUEL SET BCM_COMPTEUR=1');
end;

end.

