{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 28/01/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : UTOFPG_RAZSAL ()
Mots clefs ... : PAIE;TOF;UTOFPG_RAZSAL;PGRAZSALARIE
*****************************************************************}
{
PT1    : 03/06/2004 VG V_50 Fignolage - FQ N°11330
PT2    : 04/08/2004 VG V_50 Traitement de la table INTERIMAIRES
PT3    : 29/09/2004 VG V_50 Fonction ne marchait pas sous Oracle - FQ N°11618
PT4    : 11/08/2005 PH V_60 FQ 12101 Test si des écritures dans la compta sont
                            passées sur le compte auxiliaire
PT5    : 22/05/2007 VG V_72 Intégration du planning unifié

}
unit UTOFPG_MulRAZSAL;

interface

uses Controls,
  Classes,
{$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  mul,
  HDB,
{$ELSE}
  eMul,
{$ENDIF}
  uTob,
  sysutils,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF,
  ed_tools,
  hstatus,
  EntPaie,
  PgOutils2,
  p5Def,
  HTB97,
  YRessource;

type
  TOF_PGMULRAZSAL = class(TOF)
  private
    BCherche: TToolbarButton97;
{$IFNDEF EAGLCLIENT}
    Liste: THDBGrid;
{$ELSE}
    Liste: THGrid;
{$ENDIF}
    procedure CalculerClick(Sender: TObject);
    procedure SalarieExit(Sender: TObject);
    procedure DeleteSal(TOB_TABLE: TOB; Trace: TStringList);
  public
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnDisplay; override;
    procedure OnClose; override;
    procedure OnCancel; override;
  end;

implementation

procedure TOF_PGMULRAZSAL.OnNew;
begin
  inherited;
end;


procedure TOF_PGMULRAZSAL.OnDelete;
begin
  inherited;
end;


procedure TOF_PGMULRAZSAL.OnUpdate;
begin
  inherited;
  SetControlEnabled('BCALCULER', True);
end;


procedure TOF_PGMULRAZSAL.OnLoad;
begin
  inherited;
end;


procedure TOF_PGMULRAZSAL.OnArgument(S: string);
var
  SALAR: THEdit;
  Calculer: TToolbarButton97;
  Num: integer;
begin
  inherited;
  SALAR := THEdit(GetControl('PSA_SALARIE'));

  if SALAR <> nil then
    SALAR.OnExit := SalarieExit;

  Calculer := TToolbarButton97(GetControl('BCALCULER'));
  if Calculer <> nil then
    Calculer.OnClick := CalculerClick;

  for Num := 1 to VH_Paie.PGNbreStatOrg do
  begin
    if Num > 4 then
      Break;
    VisibiliteChampSalarie(IntToStr(Num),
      GetControl('PSA_TRAVAILN' + IntToStr(Num)),
      GetControl('TPSA_TRAVAILN' + IntToStr(Num)));
  end;
  VisibiliteStat(GetControl('PSA_CODESTAT'), GetControl('TPSA_CODESTAT'));
  BCherche := TToolbarButton97(GetControl('BCherche'));
end;


procedure TOF_PGMULRAZSAL.OnClose;
begin
  inherited;
end;


procedure TOF_PGMULRAZSAL.OnDisplay();
begin
  inherited;
end;


procedure TOF_PGMULRAZSAL.OnCancel();
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 28/01/2004
Modifié le ... :   /  /
Description .. : Procédure exécutée lors du click sur le bouton "Calculer"
Mots clefs ... : PAIE;PGRAZSALARIE
*****************************************************************}

procedure TOF_PGMULRAZSAL.CalculerClick(Sender: TObject);
var
  StTable, StUtil: string;
  i, reponse: integer;
  Maintenant: TDateTime;
  Trace: TStringList;
  Q: TQuery;
  TOB_TABLE: TOB;
  REVISEUR: Boolean;
begin
  REVISEUR := False; //PT1
{$IFNDEF EAGLCLIENT}
  Liste := THDBGrid(GetControl('FListe'));
{$ELSE}
  Liste := THGrid(GetControl('FListe'));
{$ENDIF}

  if Liste <> nil then
  begin
    if (Liste.NbSelected = 0) and (not Liste.AllSelected) then
    begin
      MessageAlerte('Aucun élément sélectionné');
      exit;
    end;

{PT1
   reponse:=PGIAsk('Vous allez initialiser toutes les tables liées aux salariés#13#10'+
                   'Voulez-vous quand-même continuer ?', Ecran.Caption);
}
    reponse := PGIAsk('Vous allez supprimer les salariés sélectionnés ainsi#13#10' +
      'que leur historique : bulletins, cumuls,#13#10' +
      'congés payés, etc.#13#10' +
      'Confirmez-vous le traitement de suppression ?',
      Ecran.Caption);
    if (reponse <> mrYes) then
      exit;

    reponse := PGIAsk('Confirmez-vous la suppression des salariés#13#10' +
      'sélectionnés et leur historique ?', Ecran.Caption);
    if (reponse <> mrYes) then
      exit;

    StUtil := 'SELECT US_CONTROLEUR' +
      ' FROM UTILISAT WHERE' +
      ' US_UTILISATEUR="' + V_PGI.FUser + '"';
    Q := OpenSql(StUtil, True);
    if not Q.EOF then
      REVISEUR := (Q.FindField('US_CONTROLEUR').AsString = 'X');
    Ferme(Q);

    if (REVISEUR = False) then
    begin
      PGIInfo('Vous n''avez pas les droits pour supprimer des salariés',
        Ecran.Caption);
      exit;
    end;
//FIN PT1

    Trace := TStringList.Create;
    StTable := 'SELECT DT_NOMTABLE, DT_PREFIXE, DH_NOMCHAMP' +
      ' FROM DETABLES' +
      ' LEFT JOIN DECHAMPS ON' +
      ' DH_PREFIXE=DT_PREFIXE WHERE' +
      ' DH_NOMCHAMP LIKE "%_SALARIE" AND' +
      ' DH_PREFIXE LIKE "P%" AND' +
      ' DT_DOMAINE="P" AND' +
      ' DT_PREFIXE<>"PCH" GROUP BY DT_NOMTABLE, DT_PREFIXE, DH_NOMCHAMP' +
      ' ORDER BY DT_NOMTABLE';
    Q := OpenSql(StTable, True);
    TOB_TABLE := TOB.CREATE('LISTE DES TABLES', nil, -1);
    TOB_TABLE.LoadDetailDB('TABLESAL', '', '', Q, True);
    Ferme(Q);

    if (Liste.AllSelected = TRUE) then
    begin
      InitMoveProgressForm(nil, 'Traitement en cours', 'Veuillez patienter SVP ...', TFmul(Ecran).Q.RecordCount, FALSE, TRUE);
      InitMove(TFmul(Ecran).Q.RecordCount, '');

      Maintenant := Now;
      if Trace <> nil then
        Trace.Add('Début du suppression à ' + TimeToStr(Maintenant));
{$IFDEF EAGLCLIENT}
      if (TFMul(Ecran).bSelectAll.Down) then
        TFMul(Ecran).Fetchlestous;
{$ENDIF}
      TFmul(Ecran).Q.First;
      while not TFmul(Ecran).Q.EOF do
      begin
        DeleteSal(TOB_TABLE, Trace);
        TFmul(Ecran).Q.Next;
      end;

      Liste.AllSelected := False;
      TFMul(Ecran).bSelectAll.Down := Liste.AllSelected;
    end
    else
    begin
      InitMoveProgressForm(nil, 'Traitement en cours', 'Veuillez patienter SVP ...', Liste.NbSelected, FALSE, TRUE);
      InitMove(Liste.NbSelected, '');

      Maintenant := Now;
      if Trace <> nil then
        Trace.Add('Début de suppression à ' + TimeToStr(Maintenant));

      for i := 0 to Liste.NbSelected - 1 do
      begin
        Liste.GotoLeBOOKMARK(i);
{$IFDEF EAGLCLIENT}
        TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
{$ENDIF}
        DeleteSal(TOB_TABLE, Trace);
      end;

      Liste.ClearSelected;
    end;

    FiniMove;
    FiniMoveProgressForm;
    PGIBox('Traitement terminé', Ecran.Caption);
    Maintenant := Now;
    if Trace <> nil then
      Trace.Add('Traitement terminé à ' + TimeToStr(Maintenant));

    CreeJnalEvt('001', '030', 'OK', nil, nil, Trace);
    if Trace <> nil then
    begin
      FreeAndNil(Trace);
    end;
    FreeAndNil(TOB_TABLE);
  end;
  if BCherche <> nil then
    BCherche.click;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 28/01/2004
Modifié le ... :   /  /
Description .. : Sortie de la zone de saisie du code salarié
Mots clefs ... : PAIE;PGRAZSALARIE
*****************************************************************}

procedure TOF_PGMULRAZSAL.SalarieExit(Sender: TObject);
begin

  if (isnumeric(GetControlText('PSA_SALARIE')) and (VH_PAIE.PgTypeNumSal = 'NUM')) then
    SetControlText('PSA_SALARIE', ColleZeroDevant(StrToInt(GetControlText('PSA_SALARIE')), 10));
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 02/02/2004
Modifié le ... :   /  /
Description .. : Suppression du salarié
Mots clefs ... : PAIE;PGRAZSALARIE
*****************************************************************}
procedure TOF_PGMULRAZSAL.DeleteSal(TOB_TABLE: TOB; Trace: TStringList);
var
CodeAuxi, NomChamp, NomTable, St, StAuxi, StSuppr: string;
Maintenant: TDateTime;
QRechAuxi: TQuery;
T_TABLED: TOB;
NBREEC : Integer;
begin
St:= TFmul (Ecran).Q.FindField ('PSA_SALARIE').asstring;
if St <> '' then
   begin
   if (isnumeric (St) and (VH_PAIE.PgTypeNumSal = 'NUM')) then
      St:= ColleZeroDevant (StrToInt(St), 10);
   try
      begintrans;
      CodeAuxi:= '';
      if Trace <> nil then
         Trace.Add ('Suppression du salarié '+St);
      StAuxi:= 'SELECT PSA_AUXILIAIRE'+
               ' FROM SALARIES WHERE'+
               ' PSA_SALARIE="'+St+'"';
      QRechAuxi:= OpenSql (StAuxi, TRUE);
      if (not QRechAuxi.EOF) then
          CodeAuxi:= QRechAuxi.FindField ('PSA_AUXILIAIRE').Asstring;
      Ferme (QRechAuxi);
//Traitement de suppression
      T_TABLED:= TOB_TABLE.FindFirst ([''], [''], False);
      while T_TABLED <> nil do
            begin
            NomTable:= T_TABLED.GetValue ('DT_NOMTABLE');
            NomChamp:= T_TABLED.GetValue ('DH_NOMCHAMP');
            if (Length (NomChamp)<=11) then
               begin
               StSuppr:= 'DELETE FROM '+NomTable+' WHERE '+
                         NomChamp+'="'+St+'"';
               ExecuteSql (StSuppr);
               end;
            T_TABLED:= TOB_TABLE.FindNext ([''], [''], False);
            end;

      StSuppr:= 'DELETE FROM PROFILSPECIAUX WHERE'+
                ' PPS_CODE="'+St+'" AND'+
                ' PPS_ETABSALARIE="-"';
      ExecuteSql (StSuppr);

      StSuppr:= 'DELETE FROM INTERIMAIRES WHERE'+
                ' PSI_INTERIMAIRE="'+St+'" AND'+
                ' PSI_TYPEINTERIM="SAL"';
      ExecuteSql (StSuppr);

      StSuppr:= 'DELETE FROM RHCOMPETRESSOURCE WHERE'+
                ' PCH_SALARIE="'+St+'" AND'+
                ' PCH_TYPERESSOURCE="SAL"';
      ExecuteSql (StSuppr);

      if (CodeAuxi <> '') then
         begin
         StSuppr:= 'SELECT COUNT (*) NBRE'+
                   ' FROM ECRITURE WHERE'+
                   ' E_AUXILIAIRE="'+CodeAuxi+'"';
         QRechAuxi:= OpenSql (StSuppr, TRUE);
         if (not QRechAuxi.EOF) then
            NBREEC:= QRechAuxi.FindField ('NBRE').AsInteger
         else
            NBREEC:= 0;
         Ferme (QRechAuxi);
         if NBREEC <= 0 then
            begin
            StSuppr:= 'DELETE FROM RIB WHERE'+
                      ' R_AUXILIAIRE="'+CodeAuxi+'"';
            ExecuteSql (StSuppr);
            StSuppr:= 'DELETE FROM TIERS WHERE'+
                      ' T_AUXILIAIRE="'+CodeAuxi+'" AND'+
                      ' T_NATUREAUXI="SAL"';
            ExecuteSql (StSuppr);
            end;
         end;
      StSuppr:= 'DELETE FROM VENTANA WHERE'+
                ' YVA_TABLEANA="PPU" AND'+
                ' YVA_IDENTIFIANT like "'+St+';%" AND'+
                ' YVA_NATUREID="PG"';
      ExecuteSql (StSuppr);

      StSuppr:= 'DELETE FROM VENTIL WHERE'+
                ' ((V_NATURE LIKE "PG%") OR'+
                ' (V_NATURE LIKE "RR%") OR'+
                ' (V_NATURE LIKE "RC%") OR'+
                ' (V_NATURE LIKE "SA%")) AND'+
                ' V_COMPTE like "'+St+'%"';
      ExecuteSql (StSuppr);

      DeleteYRS (St, '', '');   //PT5
      CommitTrans;
   except
         Rollback;
         Maintenant:= Now;
         if Trace <> nil then
            Trace.Add ('Salarié '+St+' : Suppression annulée à '+
                       TimeToStr (Maintenant));
         end;
   end;
MoveCur (False);
MoveCurProgressForm (St);
end;


initialization
  registerclasses([TOF_PGMULRAZSAL]);
end.

