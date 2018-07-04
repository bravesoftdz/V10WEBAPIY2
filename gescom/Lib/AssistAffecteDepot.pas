unit AssistAffecteDepot;


interface

uses Windows,
     Messages,
     SysUtils,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      db,
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
      FE_Main,
      HDB,
{$ENDIF}
     Classes,
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
     Mask, 
     HPop97,
     UtilUtilitaires, HPanel;

    procedure Assist_AffecteDepot;

type
  TFAssistAffecteDepot = class(TFAssist)
    TSExplications: TTabSheet;
    GBExplications: TGroupBox;
    Lib1: THLabel;
    TSConfirmation: TTabSheet;
    GBCreateDepot: TGroupBox;
    TMCodeDep: THLabel;
    DepotCode: THCritMaskEdit;
    TMLibelleDep: THLabel;
    DepotLib: THCritMaskEdit;
    Lib2: THLabel;
    Lib5: THLabel;
    Lib3: THLabel;
    LBRecap: TListBox;
    Lib4: THLabel;
    Lib6: THLabel;
    TSRapport: TTabSheet;
    LBRapport: TListBox;
    PopButton971: TPopButton97;
    bImprimer: TToolbarButton97;
    //Evènements de la forme
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure DepotCodeChange(Sender: TObject);
    procedure DepotLibChange(Sender: TObject);
    //Gestion des boutons
    procedure bSuivantClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure bImprimerClick(Sender: TObject);

  private
    procedure ActiveSuivant;
    //Execution
    procedure ExecutionAffect;
    procedure MajAffect;
    procedure MajRapportJnal(Msg : string);

  public

  end;

var
  FAssistAffecteDepot: TFAssistAffecteDepot;

const
TexteMessage: array[1..10] of string =(
          {1}  'INFORMATION',
          {2}  'Tous les mouvements de la base contiennent un dépôt.'+#13+#10+
               'L''éxécution de cet utilitaire n''est pas nécessaire.',
          {3}  'CONFIRMATION',
          {4}  'Veuillez confirmer la mise à jour des mouvements sans dépôt.',
          {5}  'ERREUR',
          {6}  'Le code du dépôt ne peut être supérieur à 3 alphanumériques.',
          {7}  'Le libellé du dépôt ne peut être supérieur à 35 alphanumériques.',
          {8}  'Ce code de dépôt existe déjà.',
          {9}  '',
          {10} ''
          );

implementation

uses Hent1,
     UTOB,
     ParamSoc,
     HStatus;

var EventsLog : TStringList;
    TobDesTables,TobDepot : TOB;
    MsgErr : string;

{$R *.DFM}

procedure Assist_AffecteDepot;
var AffDepot_Assist : TFAssistAffecteDepot;

    function TestContinue(Table : string) : boolean;
    var Qry : TQuery;
    begin
        Qry := OpenSql('SELECT COUNT(*) AS QTE FROM '+Table+' WHERE '+
                       TableToPrefixe(Table)+'_DEPOT =""', True,-1,'',true);
        if Qry.FindField('QTE').AsInteger <= 0 then
            Result := False
        else
            Result := True;
        Ferme(Qry);
    end;

begin
    if (not TestContinue('DISPO')) and
       (not TestContinue('LIGNE')) then
    begin
        HShowMessage('0;'+TraduireMemoire(TexteMessage[1])+';'+
                          TraduireMemoire(TexteMessage[2])+';W;O;O;O;','','');
        exit;
    end;
    AffDepot_Assist := TFAssistAffecteDepot.Create(Application);
    try
        AffDepot_Assist.ShowModal;
    finally
        AffDepot_Assist.free;
    end;
end;

{===============================================================================
 ============================= Evenements de la forme ==========================
 ===============================================================================}
procedure TFAssistAffecteDepot.FormShow(Sender: TObject);
begin
    inherited;
    if GetParamSoc('SO_GCDEPOTDEFAUT') = '' then
        Lib2.Visible := True
        else
        Lib2.Visible := False;
    bPrecedent.Enabled := False;
    bSuivant.Enabled := False;
    bFin.Enabled := False;
    bImprimer.Enabled := False;
    TobDesTables := TOB.Create('DETABLES', nil, -1);
    TobDepot := TOB.Create('DEPOTS', nil, -1);
    TobDepot.InitValeurs;
    EventsLog := TStringList.Create;
end;

procedure TFAssistAffecteDepot.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
    inherited;
    TobDesTables.Free;
    TobDepot.Free;
    EventsLog.Free;
    LBRecap.Free;
    LBRapport.Free;
    P.Free;
end;

procedure TFAssistAffecteDepot.DepotCodeChange(Sender: TObject);
begin
    inherited;
    ActiveSuivant;
end;

procedure TFAssistAffecteDepot.DepotLibChange(Sender: TObject);
begin
    inherited;
    ActiveSuivant;
end;

procedure TFAssistAffecteDepot.ActiveSuivant;
begin
    if (DepotCode.Text = '') or (DepotLib.Text = '') then
        bSuivant.Enabled := False
        else
        bSuivant.Enabled := True;
end;
{===============================================================================
 ============================= Gestion des boutons =============================
 ===============================================================================}
procedure TFAssistAffecteDepot.bSuivantClick(Sender: TObject);

    procedure Erreur(NumMsg : integer);
    begin
        HShowMessage('0;'+TraduireMemoire(TexteMessage[5])+';'+
                          TraduireMemoire(TexteMessage[NumMsg])+';W;A;A;A;','','');
        bSuivant.Enabled := False;
    end;

begin
    if P.ActivePage.Name = 'TSExplications' then
    begin
        if Length(DepotCode.Text) > 3 then
        begin
            Erreur(6);
            DepotCode.SetFocus;
            exit;
        end else
        if Length(DepotLib.Text) > 35 then
        begin
            Erreur(7);
            DepotLib.SetFocus;
            exit;
        end else
        begin
            if ExisteSQL('SELECT GDE_DEPOT FROM DEPOTS WHERE GDE_DEPOT="'+DepotCode.Text+'"') then
            begin
                Erreur(8);
                DepotCode.Text := '';
                DepotCode.SetFocus;
                exit;
            end;
        end;
    end;
    inherited;
    LBRecap.Clear;
    LBRecap.Items.Add('');
    LBRecap.Items.Add('- '+TraduireMemoire('Création du dépôt "')+DepotLib.Text+'"');
    LBRecap.Items.Add('');
    LBRecap.Items.Add('- '+TraduireMemoire('Affectation du dépôt "')+DepotLib.Text+'"');
    LBRecap.Items.Add('  '+TraduireMemoire('aux mouvements sans dépôt.'));
    if GetParamSoc('SO_GCDEPOTDEFAUT') = '' then
    begin
        LBRecap.Items.Add('');
        LBRecap.Items.Add('- '+TraduireMemoire('Déclaration du dépôt "')+DepotLib.Text+'"');
        LBRecap.Items.Add('  '+TraduireMemoire('comme dépôt par défaut.'));
    end;
    bSuivant.Enabled := False;
    bPrecedent.Enabled := True;
    bFin.Enabled := True;
end;

procedure TFAssistAffecteDepot.bPrecedentClick(Sender: TObject);
begin
    inherited;
    bSuivant.Enabled := True;
    bPrecedent.Enabled := False;
    bFin.Enabled := False;
end;

procedure TFAssistAffecteDepot.bFinClick(Sender: TObject);
begin
    inherited;
    if bFin.Caption = 'Fin' then
    begin
        if HShowMessage('0;'+TraduireMemoire(TexteMessage[3])+';'+
                        TraduireMemoire(TexteMessage[4])+';W;YN;N;N;','','') = mrYes then
            ExecutionAffect;
    end else
    begin
        Close;
    end;
end;

procedure TFAssistAffecteDepot.bImprimerClick(Sender: TObject);
var TobToPrint : TOB;
    Cpt : integer;
begin
  inherited;
  TobToPrint := TOB.Create('',nil,-1);
  for Cpt := 0 to LBRapport.Items.Count -1 do
    UtilTobCreat(TobToPrint,'','',LBRapport.Items[Cpt],'');
  UtilTobPrint(TobToPrint,Caption,1);
  TobToPrint.free;
end;

{===============================================================================
 ============================= Exécution =======================================
 ===============================================================================}
procedure TFAssistAffecteDepot.ExecutionAffect;
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
    InitMove(13,'');
    MoveCur(False);
    Qry := OpenSQL('SELECT * FROM DETABLES', True,-1,'',true);
    TobDesTables.LoadDetailDB('','','',Qry, False, False);
    Ferme(Qry);
    TobJnal := TOB.Create('JNALEVENT', nil, -1);
    io := Transactions(MajAffect,0);
    case io of
             oeOk : begin
                        TobJnal.PutValue('GEV_ETATEVENT', 'OK');
                        Msg := TraduireMemoire('Le traitement s''est effectué correctement.');
                        MajRapportJnal(Msg);
                    end;
        oeUnknown : begin
                        TobJnal.PutValue('GEV_ETATEVENT', 'ERR');
                        MajRapportJnal(MsgErr);
                        Msg := TraduireMemoire('*** LE TRAITEMENT A ETE ANNULE ***');
                        MajRapportJnal(Msg);
                    end;
    end;

    //Maj journal d'évènements
    MoveCur(False);
    NumEvt := 0;
    Qry := OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT', true,-1,'',true);
    if not Qry.Eof then NumEvt := Qry.Fields[0].AsInteger;
    Ferme(Qry);
    inc(NumEvt);
    TobJnal.PutValue('GEV_NUMEVENT', NumEvt);
    TobJnal.PutValue('GEV_TYPEEVENT', 'UTI');
    TobJnal.PutValue('GEV_LIBELLE', TraduireMemoire('Affectation dépôt aux mouvements'));
    TobJnal.PutValue('GEV_DATEEVENT', Date);
    TobJnal.PutValue('GEV_UTILISATEUR', V_PGI.User);
    TobJnal.PutValue('GEV_BLOCNOTE',EventsLog.Text);
    TobJnal.InsertDB(nil);
    TobJnal.Free;
    bImprimer.Enabled := True;
    FiniMove;
end;

procedure TFAssistAffecteDepot.MajAffect;
var NomTable,LibTable,Prefixe,Msg : string;
    TobTmp : TOB;
    Qte : integer;

   procedure UpdateLaTable(Table : string);
   begin
        MoveCur(False);
        TobTmp := TobDesTables.FindFirst(['DT_NOMTABLE'],[Table],True);
        if TobTmp <> nil then
        begin
            Prefixe := TableToPrefixe(Table);
            LibTable := TobTmp.GetValue('DT_LIBELLE');
            MSgErr := TraduireMemoire('*** ERREUR LORS DE LA MISE A JOUR DE "')+
                      LibTable+'" ***';
            Qte := ExecuteSql('UPDATE '+Table+' SET '+Prefixe+'_DEPOT="'+DepotCode.Text+'" '+
                              'WHERE '+Prefixe+'_DEPOT=""');
            if Qte >= 0 then
            begin
                Msg := TraduireMemoire('Mise à jour de ')+IntToStr(Qte)+' '+
                       TraduireMemoire('lignes(s) dans "')+LibTable+'".';
                MajRapportJnal(Msg);
            end;
        end;
   end;

begin
    MoveCur(False);
    Msg := TraduireMemoire('Création du dépôt "')+DepotLib.Text+'"';
    MSgErr := TraduireMemoire('*** ERREUR LORS DE LA CREATION DU DEPOT "')+DepotLib.Text+'" ***';
    TobDepot.PutValue('GDE_DEPOT', DepotCode.Text);
    TobDepot.PutValue('GDE_LIBELLE', DepotLib.Text);
    TobDepot.InsertDB(nil);
    if V_PGI.ioError <> oeOk then exit;
    MajRapportJnal(Msg);

    if GetParamSoc('SO_GCDEPOTDEFAUT') = '' then
    begin
        MoveCur(False);
        Msg := TraduireMemoire('Déclaration du dépôt "')+DepotLib.Text+
               TraduireMemoire('" comme dépôt par défaut');
        MSgErr := TraduireMemoire('*** ERREUR LORS DE LA DECLARATION DU DEPOT "')+DepotLib.Text+
                  TraduireMemoire('" COMME DEPOT PAR DEFAUT ***');
        SetParamSoc('SO_GCDEPOTDEFAUT', DepotCode.Text);
        if V_PGI.ioError <> oeOk then exit;
        MajRapportJnal(Msg);
    end;

    MoveCur(False);
    Msg := TraduireMemoire('Affectation du dépôt "')+DepotLib.Text+
           TraduireMemoire('" aux mouvements sans dépôt.');
    MajRapportJnal(Msg);
    UpdateLaTable('LIGNEOUVxx');
    UpdateLaTable('DISPO');
    UpdateLaTable('DISPOCONTREM');
    UpdateLaTable('DISPOLOT');
    UpdateLaTable('DISPOSERIE');
    UpdateLaTable('EMPLACEMENT');
    UpdateLaTable('LIGNE');
    UpdateLaTable('PIECE');
    //Inventaire - Si pas de dépôt et non validés, validation
    NomTable := 'LISTEINVENT';
    MoveCur(False);
    TobTmp := TobDesTables.FindFirst(['DT_NOMTABLE'],[NomTable],True);
    if TobTmp <> nil then
    begin
        MoveCur(False);
        Prefixe := TableToPrefixe(NomTable);
        LibTable := TobTmp.GetValue('DT_LIBELLE');
        MSgErr := TraduireMemoire('*** ERREUR LORS DE LA VALIDATION DE "')+LibTable+'" ***';
        Qte := ExecuteSql('UPDATE '+NomTable+' SET GIE_VALIDATION="X" '+
                          'WHERE GIE_DEPOT="" AND GIE_VALIDATION="-"');
        if Qte >= 0 then
        begin
            Msg := TraduireMemoire('Validation de ')+IntToStr(Qte)+' '+
                   TraduireMemoire('lignes(s) dans "')+LibTable+'".';
            MajRapportJnal(Msg);
        end;
    end;
end;

procedure TFAssistAffecteDepot.MajRapportJnal(Msg : string);
begin
    EventsLog.Add(Msg);
    LBRapport.Items.Add(Msg);
end;

end.
