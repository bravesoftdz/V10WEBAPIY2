// MVG 21/04/06 - FQ 17960 - pour syntaxe en MSACCESS
// BTY 05/06    - FQ 18119 - Positionner les indicateurs de modif de la compta
// MBO 30/10/07 - FQ 21754 - Ajout de 2 paramètres ds appel planInfo.calcul

unit AmRapport;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Vierge, Grids, Hctrls, ExtCtrls, HPanel, HSysMenu, HTB97, uTOB, StdCtrls,
  HEnt1, ed_tools,HMsgBox,
  ImEnt
{$IFDEF EAGLCLIENT}
  , UtileAGL
{$ELSE}
  , PrintDbg
  {$IFNDEF DBXPRESS} ,dbtables {$ELSE} ,uDbxDataSet {$ENDIF}
{$ENDIF}
  , Outils
  , ImPlanInfo
  , ImOuPlan
  , ImPlan
  , AMIMMO
  , AmSortie
  , IMMO_TOM
  ;

const
  COL_DATE = 3;
  COL_CORRIGE = 4;
  MAX_ACTION = 9;
  INFO_ACTION1 = 'Certaines immobilisations peuvent présenter une dotation sur l''exercice en cours '
    +
    'alors qu''elles ont été antérieurement sorties.' +
    'Lancer cet utilitaire permettra d''intégrer les immobilisations concernées dans l''historique.';
  INFO_ACTION2 =
    'Cet outil vérifie que l''action de création de l''immobilisation a bien été prise en compte.';
  INFO_ACTION3 = 'Contrôle de la valeur d''achat de l''immoblisation.';
  INFO_ACTION4 = 'Suite à une clôture, des immobilisations complètement amorties sur l''exercice précédent '
    +
    'apparaissent dans les éditions avec une VNC. Cet utilitaire permettra de mettre ' +
    'à jour le plan d''amortissement de l''immobilisation.';
  INFO_ACTION5 = 'Pour apparaître correctement dans les éditions, les immobilisations '
    +
    'de type "financière" ou "location financière" doivent être non amortissables.';
  INFO_ACTION6 =
    'Cet utilitaire contrôle la cohérence de la durée de la reprise.';
  INFO_ACTION7 =
    'Vérification des informations relatives à la gestion de l''analytique.';
  INFO_ACTION8 = 'Contrôle du plan d''amortissement.';
  INFO_ACTION9 = 'Permet de détecter les désynchronisations entre les plans,l''historique et les fiches.';

type
  PActionVerifAmort = procedure(var OkVerif: boolean; TDetail: TOB; bVerif:
    boolean);

  TActionVerifAmort = record
    NumAction: integer;
    ToDo: PActionVerifAmort;
    Info: string;
  end;

  TInfoLog = record
    TVARecuperable: double;
    TVARecuperee: double;
  end;

  TFRapportVerifAmort = class(TFVierge)
    HPanel1: THPanel;
    FListe: THGrid;
    FBCorrection: TToolbarButton97;
    CNombre: THLabel;
    BToutSel: TToolbarButton97;
    BRefresh: TToolbarButton97;
    BReinit: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure FBCorrectionClick(Sender: TObject);
    procedure FListeFlipSelection(Sender: TObject);
    procedure BToutSelClick(Sender: TObject);
    procedure BRefreshClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure BReinitClick(Sender: TObject);
    procedure FListeRowEnter(Sender: TObject; Ou: Integer;
      var Cancel: Boolean; Chg: Boolean);
  private
    { Déclarations privées }
    fNSelection: integer;
    fTitre : string;            // caption de la fenêtre
    procedure MajNombreSelection;
    procedure AfficheLaListe;
  public
    { Déclarations publiques }
    fRapport: TOB;
    fNumAction: integer;
    fListeVerifAmort: array of TActionVerifAmort;
  end;

procedure AfficheRapportVerifAmort(NumAction: integer; T: TOB; AVF: array of
  TActionVerifAmort; stTitre : string);
procedure VerifSortieAnteriorite(var OkVerif: boolean; TDetail: TOB; bVerif:
  boolean);
procedure VerifLogAcquisition(var OkVerif: boolean; TDetail: TOB; bVerif:
  boolean);
procedure ChargeListeActionVerifAmort(var AVF: array of TActionVerifAmort);
procedure VerifVNC(var OkVerif: boolean; TDetail: TOB; bVerif: boolean);
procedure VerifNamFin(var OkVerif: boolean; TDetail: TOB; bVerif: boolean);
procedure VerifDureeReprise(var OkVerif: boolean; TDetail: TOB; bVerif:
  boolean);
procedure VerifAnalytique(var OkVerif: boolean; TDetail: TOB; bVerif: boolean);
procedure VerifValAchat(var OkVerif: boolean; TDetail: TOB; bVerif: boolean);
procedure VerifPlanAMort(var OkVerif: boolean; TDetail: TOB; bVerif: boolean);
procedure VerifSynchroPlan(var OkVerif: boolean; TDetail: TOB; bVerif: boolean);
procedure ChargeListeErreurVNC(T: TOB);
function CorrigeErreurVNC(T: TOB): boolean;

implementation

{$R *.DFM}

procedure ChargeListeActionVerifAmort(var AVF: array of TActionVerifAmort);
begin
  AVF[1].NumAction := 1;
  AVF[1].ToDo := VerifSortieAnteriorite;
  AVF[1].Info := INFO_ACTION1;
  AVF[2].NumAction := 2;
  AVF[2].ToDo := VerifLogAcquisition;
  AVF[2].Info := INFO_ACTION2;
  AVF[3].NumAction := 3;
  AVF[3].ToDo := VerifValAchat;
  AVF[3].Info := INFO_ACTION3;
  AVF[4].NumAction := 4;
  AVF[4].ToDo := VerifVNC;
  AVF[4].Info := INFO_ACTION4;
  AVF[5].NumAction := 5;
  AVF[5].ToDo := VerifNamFin;
  AVF[5].Info := INFO_ACTION5;
  AVF[6].NumAction := 6;
  AVF[6].ToDo := VerifDureeReprise;
  AVF[6].Info := INFO_ACTION6;
  AVF[7].NumAction := 7;
  AVF[7].ToDo := VerifAnalytique;
  AVF[7].Info := INFO_ACTION7;
  AVF[8].NumAction := 8;
  AVF[8].ToDo := VerifPlanAmort;
  AVF[8].Info := INFO_ACTION8;
  AVF[9].NumAction := 8;
  AVF[9].ToDo := VerifSynchroPlan;
  AVF[9].Info := INFO_ACTION9;
end;

procedure VerifDureeReprise(var OkVerif: boolean; TDetail: TOB; bVerif:
  boolean);
var
  St: string;
  nEnreg: integer;
begin
  OkVerif := False;
  if bVerif then
  begin
    St := 'SELECT I_IMMO,I_LIBELLE,I_DATEPIECEA,I_PLANACTIF ';
    St := St + ' FROM IMMO ';
    St := St + ' WHERE (I_DUREEREPRISE<0) AND I_QUALIFIMMO<>"REG"';
    TDetail.LoadDetailDBFromSQL('', St);
    OkVerif := (TDetail.Detail.Count = 0);
  end
  else
  begin
    BeginTrans;
    try
      { Mise à jour de la fiche }
      nEnreg := ExecuteSQL('UPDATE IMMO SET I_DUREEREPRISE=0 WHERE I_IMMO="' +
        TDetail.GetValue('I_IMMO') + '"');
      CommitTrans;
      OkVerif := (nEnreg > 0);
    except;
      RollBack;
    end;
  end;
end;

procedure VerifNamFin(var OkVerif: boolean; TDetail: TOB; bVerif: boolean);
var
  St: string;
  nEnreg: integer;
begin
  OkVerif := False;
  if bVerif then
  begin
    St := 'SELECT I_IMMO,I_LIBELLE,I_DATEPIECEA,I_PLANACTIF ';
    St := St + ' FROM IMMO ';
    St := St +
      ' WHERE (I_NATUREIMMO="LOC" OR I_NATUREIMMO="FI") AND I_METHODEECO<>"NAM" AND I_QUALIFIMMO<>"REG"';
    TDetail.LoadDetailDBFromSQL('', St);
    OkVerif := (TDetail.Detail.Count = 0);
  end
  else
  begin
    BeginTrans;
    try
      { Mise à jour de la fiche }
      nEnreg := ExecuteSQL('UPDATE IMMO SET I_METHODEECO="NAM" WHERE I_IMMO="' +
        TDetail.GetValue('I_IMMO') + '"');
      CommitTrans;
      OkVerif := (nEnreg > 0);
    except;
      RollBack;
    end;
  end;
end;

procedure VerifVNC(var OkVerif: boolean; TDetail: TOB; bVerif: boolean);
begin
  OkVerif := False;
  if bVerif then
  begin
    ChargeListeErreurVNC(TDetail);
    OkVerif := (TDetail.Detail.Count = 0);
  end
  else
  begin
    BeginTrans;
    try
      OkVerif := CorrigeErreurVNC(TDetail);
      CommitTrans;
    except;
      RollBack;
    end;
  end;
end;

procedure VerifSortieAnteriorite(var OkVerif: boolean; TDetail: TOB; bVerif:
  boolean);
var
  Q: TQuery;
  St: string;
  nEnreg: integer;
begin
  OkVerif := False;
  if bVerif then
  begin
    St := 'SELECT  DISTINCT(I_IMMO),I_LIBELLE,I_DATEPIECEA,I_PLANACTIF ';
    St := St + ' FROM IMMOLOG ';
    St := St + ' LEFT JOIN IMMO ON I_IMMO=IL_IMMO ';
    St := St + ' WHERE IL_TYPEOP="CES" AND (I_MONTANTHT<>0  OR I_ETAT<>"FER") AND I_QUALIFIMMO<>"REG"';
    St := St + ' AND IL_DATEOP<"' + USDateTime(VHImmo^.Encours.Deb) + '"';
    Q := OpenSQL(St, True);
    OkVerif := Q.Eof;
    TDetail.LoadDetailDB('', '', '', Q, False);
    Ferme(Q);
  end
  else
  begin
    BeginTrans;
    try
      { Mise à jour du plan d'amortissement }
      ExecuteSQL('UPDATE IMMOAMOR SET IA_MONTANTECO=0, IA_MONTANTFISCAL=0, IA_BASEDEBEXOECO=0, IA_BASEDEBEXOFISC=0 WHERE IA_IMMO="'
        + TDetail.GetValue('I_IMMO') + '" AND IA_NUMEROSEQ=' +
        IntToStr(TDetail.GetValue('I_PLANACTIF')));
      { Montant HT forcé à 0 - Mise en historique}
      nEnreg :=
        ExecuteSQL('UPDATE IMMO SET I_ETAT="FER", I_MONTANTHT=0 WHERE I_IMMO="'
        +
        TDetail.GetValue('I_IMMO') + '"');
      CommitTrans;
      OkVerif := (nEnreg > 0);
    except;
      RollBack;
    end;
  end;
end;

procedure VerifLogAcquisition(var OkVerif: boolean; TDetail: TOB; bVerif:
  boolean);

  procedure EnregLogAcquisition(CodeImmo: string; DateOpe: TDateTime; PlanActif:
    integer; infos: TInfoLog);
  var
    TLog: TOB;
  begin
    // Création nouvel enregistrement ImmoLog
    TLog := TOB.Create('IMMOLOG', nil, -1);
    try
      TLog.PutValue('IL_TYPEOP', 'ACQ');
      TLog.PutValue('IL_IMMO', CodeImmo);
      TLog.PutValue('IL_DATEOP', DateOpe);
      TLog.PutValue('IL_ORDRE', 1);
      TLog.PutValue('IL_ORDRESERIE', 1);
      TLog.PutValue('IL_PLANACTIFAV', 0);
      TLog.PutValue('IL_PLANACTIFAP', PlanActif);
      TLog.PutValue('IL_LIBELLE', RechDom('TIOPEAMOR', 'ACQ', FALSE) + ' ' +
        DateToStr(DateOpe));
      TLog.PutValue('IL_TYPEMODIF', AffecteCommentaireOperation('ACQ'));
      TLog.PutValue('IL_TVARECUPERABLE', Infos.TVARecuperable);
      TLog.PutValue('IL_TVARECUPEREE', Infos.TVARecuperee);
      TLog.InsertDB(nil);
    finally
      TLog.Free;
    end;
  end;

var
  Q: TQuery;
  St: string;
  Infos: TInfoLog;

begin
  OkVerif := False;
  if bVerif then
  begin
    St :=
      'SELECT I_IMMO,I_LIBELLE,I_DATEPIECEA,I_TVARECUPERABLE,I_TVARECUPEREE,I_PLANACTIF FROM IMMO ';
    St := St +
      ' WHERE NOT EXISTS (SELECT IL_IMMO FROM IMMOLOG WHERE IL_IMMO=IMMO.I_IMMO ';
    St := St + ' AND (IL_TYPEOP="ACQ")) AND I_QUALIFIMMO<>"REG"';
    Q := OpenSQL(St, True);
    OkVerif := Q.Eof;
    TDetail.LoadDetailDB('', '', '', Q, False);
    Ferme(Q);
  end
  else
  begin
    BeginTrans;
    try
      Infos.TVARecuperable := TDetail.GetValue('I_TVARECUPERABLE');
      Infos.TVARecuperee := TDetail.GetValue('I_TVARECUPEREE');
      EnregLogAcquisition(TDetail.GetValue('I_IMMO'),
        TDetail.GetValue('I_DATEPIECEA'), TDetail.GetValue('I_PLANACTIF'),
        infos);
      CommitTrans;
      OkVerif := True;
    except
      RollBack;
    end;
  end;
end;

procedure AfficheRapportVerifAmort(NumAction: integer; T: TOB; AVF: array of
  TActionVerifAmort; stTitre : string);
var
  FRapportVerifAmort: TFRapportVerifAmort;
begin
  FRapportVerifAmort := TFRapportVerifAmort.Create(Application);
  try
    FRapportVerifAmort.fRapport := T;
    FRapportVerifAmort.fNumAction := NumAction;
    FRapportVerifAmort.fTitre := stTitre;
    FRapportVerifAmort.ShowModal;
  finally
    FRapportVerifAmort.Free;
  end;
end;

procedure VerifAnalytique(var OkVerif: boolean; TDetail: TOB; bVerif: boolean);
var
  Q: TQuery;
  St: string;
  nEnreg: integer;
begin
  OkVerif := False;
  if bVerif then
  begin
    St := 'SELECT I_IMMO,I_LIBELLE,I_DATEPIECEA,I_PLANACTIF FROM IMMO ';
    St := St +
      ' WHERE ((I_VENTILABLE="X" and I_VENTILABLE1<>"X" and I_VENTILABLE2<>"X" and I_VENTILABLE3<>"X" and I_VENTILABLE4<>"X" and I_VENTILABLE5<>"X")';
    St := St +
      ' OR (I_VENTILABLE="-" and I_VENTILABLE1<>"-" and I_VENTILABLE2<>"-" and I_VENTILABLE3<>"-" and I_VENTILABLE4<>"-" and I_VENTILABLE5<>"-")) AND I_QUALIFIMMO<>"REG"';
    Q := OpenSQL(St, True);
    OkVerif := Q.Eof;
    TDetail.LoadDetailDB('', '', '', Q, False);
    Ferme(Q);
  end
  else
  begin
    BeginTrans;
    try
      { Mise à jour du plan d'amortissement }
      nEnreg :=
        ExecuteSQL('UPDATE IMMO SET I_VENTILABLE1="-",I_VENTILABLE2="-",I_VENTILABLE3="-",I_VENTILABLE4="-",I_VENTILABLE5="-" WHERE I_IMMO="'
        + TDetail.GetValue('I_IMMO') + '" AND I_VENTILABLE="-"');
      nEnreg := nEnreg +
        ExecuteSQL('UPDATE IMMO SET I_VENTILABLE1="X",I_VENTILABLE2="X",I_VENTILABLE3="X",I_VENTILABLE4="X",I_VENTILABLE5="X" WHERE I_IMMO="'
        + TDetail.GetValue('I_IMMO') + '" AND I_VENTILABLE="X"');
      CommitTrans;
      OkVerif := (nEnreg > 0);
    except;
      RollBack;
    end;
  end;
end;

procedure VerifValAchat(var OkVerif: boolean; TDetail: TOB; bVerif: boolean);

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 16/03/2005
Modifié le ... :   /  /
Description .. : Contrôle des fiches
Suite ........ : FQ 15472 - Ne pas tenir compte des locations longue durée
Suite ........ : lors du contrôle
Mots clefs ... :
*****************************************************************}
  function _GetValeurAchat(CodeImmo: string): double;
  var
    Q: TQuery;
  begin
    Result := 0;
    Q :=
      OpenSQL('SELECT I_IMMO,I_MONTANTHT,I_VALEURACHAT,IL_VOCEDEE,IL_TYPEOP FROM IMMO LEFT JOIN IMMOLOG ON IL_IMMO=I_IMMO AND IL_PLANACTIFAP=I_PLANACTIF WHERE I_IMMO="'
      + CodeImmo + '" ORDER BY I_IMMO', True);
    try
      if not Q.Eof then
      begin
        if (Q.FindField('IL_TYPEOP').AsString = 'CES') then
          Result := Q.FindField('IL_VOCEDEE').AsFloat
        else
          Result := Q.FindField('I_MONTANTHT').AsFloat;
      end;
    finally
      Ferme(Q);
    end;
  end;

var
  Q: TQuery;
  St: string;
  nEnreg: integer;
  ValAchat : double;
begin
  OkVerif := False;
  if bVerif then
  begin
    St := 'SELECT I_IMMO,I_LIBELLE,I_DATEPIECEA,I_PLANACTIF FROM IMMO ';
    St := St +
      ' WHERE I_NATUREIMMO<>"LOC" AND (I_VALEURACHAT=0 OR'+
      ' ((I_VALEURACHAT<>I_MONTANTHT+I_TVARECUPERABLE-I_TVARECUPEREE) AND I_MONTANTHT<>0))'+
      ' AND I_ETAT<>"FER" AND I_QUALIFIMMO<>"REG"';
    Q := OpenSQL(St, True);
    OkVerif := Q.Eof;
    TDetail.LoadDetailDB('', '', '', Q, False);
    Ferme(Q);
  end
  else
  begin
    BeginTrans;
    try
      { Mise à jour de la valeur d'achat }
      ValAchat :=  _GetValeurAchat(TDetail.GetValue('I_IMMO'));
      if (ValAchat <> 0) then
      begin
        nEnreg := ExecuteSQL('UPDATE IMMO SET I_VALEURACHAT=' +
          StrfPoint(ValAchat)
          + ' WHERE I_IMMO="' + TDetail.GetValue('I_IMMO') + '"');
      end else nEnreg := 0;
      CommitTrans;
      OkVerif := (nEnreg > 0);
    except;
      RollBack;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 16/03/2005
Modifié le ... : 16/03/2005
Description .. : FQ 15269 - 16/03/2005 - CA : Ajout condition
Suite ........ : I_METHODEECO<>"NAM" pour ne pas contrôler le plan
Suite ........ : sur ces immobilisations
Suite ........ : FQ 15468 - 16/03/2005 - CA - Ajout condition
Suite ........ : I_REPRISEECO<>I_BASEECO pour ne pas traiter les cas
Suite ........ : des immobilisations totalement amorties
Mots clefs ... :
*****************************************************************}

procedure VerifPlanAmort(var OkVerif: boolean; TDetail: TOB; bVerif: boolean);
var
  Q: TQuery;
  St: string;
  Plan: TPlanAmort;
  TImmo, TPlan: TOB;
begin
  OkVerif := False;
  if bVerif then
  begin
    St := 'SELECT I_IMMO,I_LIBELLE,I_DATEPIECEA,I_PLANACTIF FROM IMMO ';
    St := St +
      ' LEFT JOIN IMMOAMOR ON (IA_IMMO=I_IMMO AND IA_NUMEROSEQ=I_PLANACTIF) ' +
        ' WHERE IA_IMMO IS NULL AND I_MONTANTHT<>0 AND (I_NATUREIMMO="PRO" OR I_NATUREIMMO="CB") '+
        ' AND I_METHODEECO<>"NAM" AND I_REPRISEECO<>I_BASEECO AND I_PLANACTIF=1 AND I_ETAT<>"FER" AND I_QUALIFIMMO<>"REG"';
    Q := OpenSQL(St, True);
    OkVerif := Q.Eof;
    TDetail.LoadDetailDB('', '', '', Q, False);
    Ferme(Q);
  end
  else
  begin
    BeginTrans;
    try
      { Recalcul du plan }
      Plan := TPlanAmort.Create(true);
      TImmo := TOB.Create('IMMO', nil, -1);
      TImmo.SelectDB('"' + TDetail.GetValue('I_IMMO') + '"', nil);
      TImmo.PutValue('I_PLANACTIF', TImmo.GetValue('I_PLANACTIF') - 1);
      Plan.CalculTOB(TImmo, iDate1900);
      TPlan := TOB.Create('', TImmo, -1);
      Plan.SauveTOB(TPlan);
      TPlan.InsertDB(nil);
      TImmo.Free;
      Plan.Free;
      CommitTrans;
      OkVerif := True;
    except;
      RollBack;
    end;
  end;
end;

procedure TFRapportVerifAmort.FormShow(Sender: TObject);
begin
  inherited;
  Caption := fTitre;
  BReinit.Enabled := False;
  FListe.ColTypes[COL_CORRIGE] := 'B';
  FListe.ColFormats[COL_CORRIGE] := IntToStr(Ord(csCheckbox));
  FListe.ColAligns[COL_DATE] := taCenter;
  FListe.ColAligns[COL_CORRIGE] := taCenter;
  SetLength(fListeVerifAmort, MAX_ACTION + 1);
  ChargeListeActionVerifAmort(fListeVerifAmort);
  AfficheLaListe;
end;

procedure TFRapportVerifAmort.FBCorrectionClick(Sender: TObject);
var
  i: integer;
  OkVerif: boolean;
  bTrouve: boolean;
begin
  inherited;
  bTrouve := False;
  { Correction des erreurs }
  InitMoveProgressForm(nil, 'Correction en cours...Veuillez patienter',
    'Correction en cours', fNSelection + 1, True, True);
  for i := 0 to fRapport.Detail.Count - 1 do
  begin
    if (FListe.IsSelected(i + 1) or FListe.AllSelected) then
    begin
      if not MoveCurProgressForm('Immobilisation n°' +
        fRapport.Detail[i].GetValue('I_IMMO')) then
        break;
      fListeVerifAmort[fNumAction].ToDo(OkVerif, fRapport.Detail[i], False);
      if OkVerif then
        // FQ 18119 Positionner les indicateurs de modif
        // FListe.Cells[COL_CORRIGE, i + 1] := 'X'
        begin
        FListe.Cells[COL_CORRIGE, i + 1] := 'X';
        VHImmo^.ChargeOBImmo := True;
        ImMarquerPublifi(True);
        end
      else
        FListe.Cells[COL_CORRIGE, i + 1] := '-';
      bTrouve := True;
      if not FListe.AllSelected then
        FListe.FlipSelection(i + 1);
    end;
  end;
  if FListe.AllSelected then
    FListe.AllSelected := False;
  if bTrouve then
    FBCorrection.Enabled := False;
  FiniMoveProgressForm;
end;

procedure TFRapportVerifAmort.FListeFlipSelection(Sender: TObject);
begin
  inherited;
  MajNombreSelection;
  BReinit.Enabled := ((FListe.Cells[COL_CORRIGE, FListe.Row] = '-') and (fNSelection=1));
end;

procedure TFRapportVerifAmort.MajNombreSelection;
var
  St: string;
  i: integer;
begin
  fNSelection := 0;
  for i := 0 to fRapport.Detail.Count - 1 do
    if FListe.IsSelected(i + 1) then
      Inc(fNSelection, 1);
  if fNSelection <= 1 then
    St := 'ligne sélectionnée'
  else
    St := 'lignes sélectionnées';
  CNombre.Caption := IntToStr(fNSelection) + '/' +
    IntToStr(fRapport.Detail.Count) + ' ' + TraduireMemoire(St);
end;

procedure TFRapportVerifAmort.BToutSelClick(Sender: TObject);
begin
  inherited;
  FListe.AllSelected := not FListe.AllSelected;
end;

procedure TFRapportVerifAmort.BRefreshClick(Sender: TObject);
var
  OkVerif: boolean;
begin
  inherited;
  fRapport.ClearDetail;
  fListeVerifAmort[fNumAction].ToDo(OkVerif, fRapport, True);
  if OkVerif then
  begin
    FListe.VidePile(False);
    MajNombreSelection;
  end
  else
    AfficheLaListe;
  FBCorrection.Enabled := True;
end;

procedure TFRapportVerifAmort.AfficheLaListe;
begin
  FListe.VidePile(False);
  fRapport.PutGridDetail(FListe, False, False,
    ';I_IMMO;I_LIBELLE;I_DATEPIECEA;;');
  FListe.ColWidths[0] := 20;
  FListe.ColWidths[1] := 100;
  FListe.ColWidths[2] := 300;
  FListe.ColWidths[3] := 100;
  FListe.ColWidths[4] := 60;
  MajNombreSelection;
end;

function CorrigeErreurVNC(T: TOB): boolean;
var
  PlanInfo: TPlanInfo;
  DateFinAmortCalc, DateDebutEx, DateFinEx: TDateTime;
  Dotation: double;
begin
  Result := False;
  Dotation := 0;
  if
    ExisteSQL('SELECT I_IMMO,I_LIBELLE,IA_DATE,IA_MONTANTECO,I_PLANACTIF FROM IMMO LEFT JOIN IMMOAMOR ON IA_IMMO=I_IMMO AND IA_NUMEROSEQ=I_PLANACTIF WHERE IA_DATE="'
    + USDateTime(VHImmo^.Encours.Fin) + '" AND IA_MONTANTECO<>0 AND I_IMMO="' +
    T.GetValue('I_IMMO') + '"') then
  begin
    PlanInfo := TPlanInfo.Create('');
    try
      PlanInfo.ChargeImmo(T.GetValue('I_IMMO'));
      PlanInfo.Plan.CalculDateFinAmortissement(PlanInfo.Plan.AmortEco);
      DateFinAmortCalc := PlanInfo.Plan.GetDateFinAmort;
      DateFinEx := iDate1900;
      GetDatesExercice(DateFinAmortCalc, DateDebutEx, DateFinEx);
      (* Si la plus grande date du dernier plan trouvé dans la base est supérieure à la
          date de fin d'exercice concerné par la fin de l'amortissement calculé, il y a problème ! *)
      if (VHImmo^.Encours.Fin > DateFinEx) then
      begin
        { Mise à jour de la dotation depuis IMMOAMOR }
        { Suppression des lignes du plan postérieures à la date de fin d'amortissement }
        ExecuteSQL('DELETE FROM IMMOAMOR WHERE IA_IMMO="' + T.GetValue('I_IMMO')
          +
          '" AND IA_NUMEROSEQ=' + IntToStr(T.GetValue('I_PLANACTIF')) +
          ' AND IA_DATE >"' + USDateTime(DateFinEx) + '"');
        { Mise à jour de la dotation pour le dernier exercice de plan }
        Result := (ExecuteSQL('UPDATE IMMOAMOR SET IA_MONTANTECO=' +
          FloatToStr(Dotation) +
          ' WHERE IA_IMMO="' + T.GetValue('I_IMMO') +
          '" AND IA_NUMEROSEQ=' + IntToStr(T.GetValue('I_PLANACTIF')) +
          ' AND IA_DATE ="' + USDateTime(DateFinEx) + '"') > 0);
      end;
    finally
      PlanInfo.Free;
    end;
  end;
end;

procedure ChargeListeErreurVNC(T: TOB);
var
  Q: TQuery;
  PlanInfo: TPlanInfo;
begin

  { ATTENTION : prendre aussi en compte le cas ou la dotation est nulle est la base est différente de 0 }

  PlanInfo := TPlanInfo.Create('');
  Q :=
    OpenSQL('SELECT I_IMMO,I_LIBELLE,I_PLANACTIF,I_VALEURACHAT,I_NATUREIMMO FROM IMMO WHERE I_ETAT<>"FER" AND I_QUALIFIMMO<>"REG" ORDER BY I_IMMO', True);
  while not Q.Eof do
  begin
    { FQ 15467 - on ne prend pas en compte les immobilisations avec un contrôle de fiche (valeur achat) incorrect. }
    if ((Q.FindField('I_VALEURACHAT').AsFloat = 0) and
      (Q.FindField('I_NATUREIMMO').AsString <> 'LOC')) then
    begin
      Q.Next;
      continue;
    end;
    PlanInfo.ChargeImmo(Q.FindField('I_IMMO').AsString);
    PlanInfo.Calcul(VHImmo^.Encours.Fin, true, false, ''); //fq 21754 - mbo - 30/10/2007
    if PlanInfo.VNCEco < 0 then
    begin
      with TOB.Create('', T, -1) do
      begin
        AddChampSupValeur('I_IMMO', Q.FindField('I_IMMO').AsString);
        AddChampSupValeur('I_LIBELLE', PlanInfo.Plan.LibelleImmo);
        AddChampSupValeur('I_DATEPIECEA', PlanInfo.Plan.DateAchat);
        AddChampSupValeur('I_PLANACTIF', Q.FindField('I_PLANACTIF').AsInteger);
      end;
    end;
    Q.Next;
  end;
  Ferme(Q);
  PlanInfo.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 16/03/2005
Modifié le ... :   /  /    
Description .. : Vérification de la cohérence du dernier plan entre les fiches, 
Suite ........ : plan d'amortissement et historique.
Mots clefs ... :
*****************************************************************}
procedure VerifSynchroPlan(var OkVerif: boolean; TDetail: TOB; bVerif: boolean);
var
  Q: TQuery;
  St: string;
  LaSortie : TAmSortie;
begin
  OkVerif := True;
  if bVerif then
  begin
    // MVG 21/04/2006 pour syntaxe en MSACCESS
    St := 'select i_immo,i_opecession,i_libelle,i_datepiecea,max(i_planactif) AS PLAN_ACTIF,max(ia_numeroseq) AS PLAN_AMOR, max(il_planactifap) AS N_ORDRE';
    St := St + ' from immo left join immolog on il_immo=i_immo ';
    St := St + ' left join immoamor on ia_immo=i_immo ';
    St := St + ' where i_etat<>"FER" and i_methodeeco<>"NAM" and i_qualifimmo<>"REG" ';
    St := St + ' group by i_immo,i_opecession,i_libelle,i_datepiecea,ia_immo,i_repriseeco,i_baseeco,i_repcedeco,i_opecession ';
    St := St + ' having (((max (ia_numeroseq)<>max(i_planactif)) or (max (il_planactifap)<>max(i_planactif))) and max(ia_montanteco)>0) ';
    St := St + ' or ((ia_immo is null) and (i_repriseeco+i_repcedeco<>i_baseeco) and (i_opecession<>"X")) ';
    St := St + ' order by i_immo';
    Q := OpenSQL(St, True);
    if not Q.Eof then
    begin
      Q.First;
      while not Q.Eof do
      begin
        OkVerif := False;
        with TOB.Create('', TDetail, -1) do
        begin
          AddChampSupValeur('I_IMMO', Q.FindField('I_IMMO').AsString);
          AddChampSupValeur('I_LIBELLE', Q.FindField('I_LIBELLE').AsString);
          AddChampSupValeur('I_DATEPIECEA', Q.FindField('I_DATEPIECEA').AsDateTime);
          // MVG 21/04/2006 pour syntaxe en MSACCESS
          AddChampSupValeur('PLAN_ACTIF', Q.FindField('PLAN_ACTIF').AsInteger);
        end;
        Q.Next;
      end;
    end;
    Ferme(Q);
  end
  else
  begin
    LaSortie := TAmSortie.Create (TDetail.GetValue('I_IMMO'));
    if LaSortie.EstSortie then
    begin
      { Annulation de la sortie}
      LaSortie.Execute;
      { Sortie }
      OkVerif := LaSortie.Execute;
    end else OkVerif := False;
  end;
end;

procedure TFRapportVerifAmort.BImprimerClick(Sender: TObject);
begin
  inherited;
{$IFNDEF EAGLCLIENT}
  PrintDBGrid(FListe, nil, Caption, '');
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 20/12/2004
Modifié le ... :   /  /
Description .. : Réinitialisation de l'immobilisation
Mots clefs ... :
*****************************************************************}

procedure TFRapportVerifAmort.BReinitClick(Sender: TObject);
var i : integer;
begin
  inherited;
  if (fNSelection <> 1) then
  begin
    PGIInfo('Veuillez ne sélectionner qu''une seule immobilisation à réinitialiser.');
    exit;
  end;
  { Correction des erreurs }
  InitMoveProgressForm(nil, 'Réinitialisation en cours...Veuillez patienter',
    'Réinitialisation en cours', fNSelection + 1, True, True);
  for i := 0 to fRapport.Detail.Count - 1 do
  begin
    if ((FListe.IsSelected(i + 1)) or (FListe.AllSelected and (FListe.RowCount=2))) then
    begin
      if not MoveCurProgressForm('Immobilisation n°' +
        fRapport.Detail[i].GetValue('I_IMMO')) then
        break;
     if not ReinitImmo(fRapport.Detail[i].GetString('I_IMMO')) then
       AMLanceFiche_FicheImmobilisation(fRapport.Detail[FListe.Row - 1].GetValue('I_IMMO'), taModif, '') //;

     // FQ 18119 Positionner les indicateurs de modif
     else
       begin
       VHImmo^.ChargeOBImmo := True;
       ImMarquerPublifi (True);
       end;

     FListe.Cells[COL_CORRIGE, i + 1] := 'X';
     if not FListe.AllSelected then FListe.FlipSelection(i + 1);
     BReinit.Enabled := False;
    end;
  end;
  if FListe.AllSelected then
    FListe.AllSelected := False;
  FiniMoveProgressForm;
end;

procedure TFRapportVerifAmort.FListeRowEnter(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
  inherited;
//  BReinit.Enabled := (FListe.Cells[COL_CORRIGE, FListe.Row] = '-') and ;
end;

end.

