unit ImMigEur;

interface

uses Controls,hctrls, ImEnt, ImPlan, SysUtils,
     Outils,HMsgBox,HEnt1,HStatus,

{$IFDEF eAGLClient}
   Utob,
{$ELSE}
   DBTables,
{$ENDIF eAGLClient}

{$IFDEF SERIE1}
{$ELSE}
     ENT1, UtilPGI,
{$ENDIF SERIE1}
     ParamSoc ;

procedure LanceMigrationImmoEuro;
function LanceMigrationImmoEuroS1: boolean ;

//procedure RecalculDernierPlanEuro (QImmo : TQuery);
procedure MajFicheImmoEuro ( var Q : TQuery; plan_encours : integer);
procedure MajPlanEcheEuro (var Q : TQuery);
function  MajImmoLogEuro ( var Q: TQuery; bInc : boolean) : integer;
procedure MajPlanAmorEuro ( var Q : TQuery; bInc : boolean);
procedure MajChampFicheImmoEuro( var Q : TQuery; sChamp, sChampContre : string);
procedure MajChampImmoEuro ( var Q : TQuery; sChamp : string);

implementation

type
  TImmoMigrationEuro = class
    procedure Execute ;
  private
    procedure ImMigrationEuro;
  end;

function LanceMigrationImmoEuroS1: boolean ;
var LaMigration : TImmoMigrationEuro;
begin
  result:=true ; //ATTENTION a changer !!!
{$IFDEF EURO}
  If VH^.VNUMSOC>500 Then
{$ENDIF}
  if GetParamSoc('SO_IMMOMIGEURO') then  exit;
  LaMigration := TImmoMigrationEuro.Create;
  try
    LaMigration.ImMigrationEuro ;
    LaMigration.Free;
  except
    LaMigration.Free;
    result:=false ;
  end ;
end;

procedure LanceMigrationImmoEuro;
var LaMigration : TImmoMigrationEuro;
    varIMMOMIGEURO : variant;
begin
  if VHImmo^.TenueEuro then
  begin
    PGIBox('Le dossier est en Euro, la migration est impossible.','Migration vers l''euro');
    exit;
  end;
{$IFDEF EURO}
  If VH^.VNUMSOC>500 Then
    BEGIN
{$ENDIF}
    varIMMOMIGEURO := GetParamSoc('SO_IMMOMIGEURO');
    if varIMMOMIGEURO then
      begin
      PGIBox('La migration vers l''euro a déjà été réalisée.','Migration vers l''euro');
      exit;
      end;
{$IFDEF EURO}
    END ;
{$ENDIF}
  if Not ImBlocageMonoPoste(True) then Exit ;
  LaMigration := TImmoMigrationEuro.Create;
  LaMigration.Execute();
  LaMigration.Free;
  ChargeVHImmo ;
  ImDeblocageMonoPoste(True) ;
end;

procedure TImmoMigrationEuro.Execute ;
begin
  if (PGIAsk('Voulez-vous migrer les données des immobilisations vers l''euro ?','Migration vers l''euro')=mrYes) then
  begin
    if (PGIAsk('Confirmez-vous la migration des données des immobilisations vers l''euro ?','Migration vers l''euro')=mrYes) then
    begin
      if Transactions(ImMigrationEuro, 1)<>oeOK then
      begin
        PGIBox('Problème lors de la migration des données en euro.','Migration vers l''euro');
      end else
      begin
{$IFDEF EURO}
      If VH^.VNUMSOC>500 Then
        BEGIN
{$ENDIF}
        SetParamSoc('SO_IMMOMIGEURO','X');
{$IFDEF EURO}
        END ;
{$ENDIF}
      PGIInfo('La migration des données en euro a été effectuée correctement.#10#13Vous devez maintenant effectuer la migration de la comptabilité.','Migration vers l''euro');
      end;
    end;
  end;
end;

procedure TImmoMigrationEuro.ImMigrationEuro;
var QImmo, QLog, QPlan, QW : TQuery; plan_ap , plan_encours, i, ordre : integer;
    code_immo : string; bImmoAnnee : boolean; Plan: TPlanAmort ;
begin
  QImmo := OpenSQL ('SELECT * FROM IMMO', False);
  InitMove (QImmo.RecordCount,'');
  while not QImmo.Eof do
  begin
    plan_encours := -1;
    code_immo := QImmo.FindField ('I_IMMO').AsString;
    MajPlanEcheEuro (QImmo);
    QLog := OpenSQL ('SELECT * FROM IMMOLOG WHERE IL_IMMO="'+code_immo+
                '" AND IL_DATEOP>="'+USDateTime(VHImmo^.Encours.Deb)+'" ORDER BY IL_DATEOP DESC', FALSE);
    if not QLog.Eof then
    begin
      bImmoAnnee := (QImmo.FindField ('I_DATEAMORT').AsDateTime >= VHImmo^.Encours.Deb);
      if bImmoAnnee then
      begin
        while not QLog.Eof do
        begin
          MajImmoLogEuro (QLog,FALSE);
          QLog.Next;
        end;
        QPlan := OpenSQL ('SELECT * FROM IMMOAMOR WHERE IA_IMMO="'+code_immo+'"',False);
        while not QPlan.Eof do
        begin
           MajPlanAmorEuro (QPlan,FALSE);
           QPlan.Next;
        end;
        Ferme (QPlan);
      end else
      begin
        plan_encours := -1;
        plan_ap := QImmo.FindField ('I_PLANACTIF').AsInteger+1;
        while not QLog.Eof do
        begin
          plan_ap := MajImmoLogEuro (QLog,TRUE);
          if plan_encours = -1 then plan_encours := plan_ap + 1;
          QPlan := OpenSQL ('SELECT * FROM IMMOAMOR WHERE IA_IMMO="'+code_immo+
                        '" AND IA_NUMEROSEQ='+IntToStr(plan_ap),False);
          while not QPlan.Eof do
          begin
            MajPlanAmorEuro (QPlan,TRUE);
            QPlan.Next;
          end;
          Ferme (QPlan);
          QLog.Next;
        end;
        QW:=OpenSQL('SELECT * FROM IMMOAMOR WHERE IA_IMMO="'+W_W+'"',False);
        QPlan := OpenSQL ('SELECT * FROM IMMOAMOR WHERE IA_IMMO="'+code_immo+
                       '" AND IA_NUMEROSEQ='+IntToStr(plan_ap-1),False);
        while not QPlan.Eof do
        begin
          QW.Insert;InitNew (QW);
          for i:=0 to QPlan.FieldCount-1 do QW.Fields[i].AsVariant:=QPlan.Fields[i].AsVariant ;
          QW.FindField ('IA_NUMEROSEQ').AsInteger := plan_ap;
          QW.FindField ('IA_CHANGEAMOR').AsString := 'EUR';
          QW.Post;
          MajPlanAmorEuro (QW,False);
          QPlan.Next
        end;
        Ferme (QW);
        Ferme (QPlan);
      end;
    end else
    begin
      plan_ap := QImmo.FindField ('I_PLANACTIF').AsInteger+1;
      QW:=OpenSQL('SELECT * FROM IMMOAMOR WHERE IA_IMMO="'+W_W+'"',False);
      QPlan := OpenSQL ('SELECT * FROM IMMOAMOR WHERE IA_IMMO="'+code_immo+
                       '" AND IA_NUMEROSEQ='+IntToStr(plan_ap-1),False);
      while not QPlan.Eof do
      begin
        QW.Insert;InitNew (QW);
        for i:=0 to QPlan.FieldCount-1 do QW.Fields[i].AsVariant:=QPlan.Fields[i].AsVariant ;
        QW.FindField ('IA_NUMEROSEQ').AsInteger := plan_ap;
        QW.FindField ('IA_CHANGEAMOR').AsString := 'EUR';
        QW.Post;
        MajPlanAmorEuro (QW,False);
        QPlan.Next
      end;
      Ferme (QW);
      Ferme (QPlan);
      plan_encours := plan_ap;
    end;
    Ferme (QLog);
    QLog := OpenSQL ('SELECT IL_ORDRE,IL_PLANACTIFAP FROM IMMOLOG WHERE IL_IMMO="'+code_immo+
                '" AND IL_DATEOP<"'+USDateTime(VHImmo^.Encours.Deb)+'" ORDER BY IL_DATEOP DESC', TRUE);
    if not QLog.Eof then
    begin
      ordre := QLog.FindField('IL_ORDRE').AsInteger;
      plan_ap := QLog.FindField('IL_PLANACTIFAP').AsInteger;
      Ferme (QLog);
      QLog := OpenSQL ('SELECT * FROM IMMOLOG WHERE IL_IMMO="'+W_W+'"',False);
      QLog.Insert;
      InitNew(QLog);
      QLog.FindField('IL_ORDRE').AsInteger := ordre + 1;
      ordre := TrouveNumeroOrdreSerieLogSuivant;
      QLog.FindField('IL_ORDRESERIE').AsInteger := ordre;
      QLog.FindField('IL_IMMO').AsString := code_immo;
      QLog.FindField('IL_PLANACTIFAV').AsFloat:=plan_ap ;
      QLog.FindField('IL_PLANACTIFAP').AsFloat:=plan_encours ;
      QLog.FindField('IL_TYPEOP').AsString := 'EUR';
      QLog.FindField('IL_LIBELLE').AsString:=RechDom('TIOPEAMOR', 'EUR', FALSE)+' '+DateToStr(VHImmo^.Encours.Deb);
      QLog.FindField('IL_TYPEMODIF').AsString:=AffecteCommentaireOperation('EUR');
      QLog.FindField('IL_DATEOP').AsDateTime := VHImmo^.Encours.Deb;
      QLog.Post;
      Ferme (QLog);
    end else Ferme (QLog);
    QImmo.Edit;
    MajFicheImmoEuro (QImmo, plan_encours);
    //RecalculDernierPlanEuro (QImmo);
    QImmo.Post;
    QImmo.Next;
    MoveCur(False);
  end;
  FiniMove;
  Ferme (QImmo);
  //lancement du recalcul des plans pour eviter les pbs d'arrondi
  Plan:=TPlanAmort.Create(true) ;
  Plan.UpdateBasePlan ;
  Plan.free;
end;

procedure MajChampFicheImmoEuro( var Q : TQuery; sChamp, sChampContre : string);
var MontantTmp : double;
begin
  if (sChamp = sChampContre) or (Q.FindField ('I_SAISIECONTRE').AsString = '-') then
  begin
    Q.FindField(sChamp).AsFloat := ImPivotToEuro(Q.FindField(sChamp).AsFloat);
  end else
  begin
    MontantTmp := Q.FindField(sChampContre).AsFloat;
    Q.FindField(sChampContre).AsFloat := Q.FindField(sChamp).AsFloat;
    Q.FindField(sChamp).AsFloat := MontantTmp;
  end;
end;

procedure MajFicheImmoEuro ( var Q : TQuery; plan_encours : integer);
begin
   MajChampFicheImmoEuro( Q,'I_MONTANTHT','I_MONTANTHTCONTRE');
   MajChampFicheImmoEuro( Q,'I_TVARECUPERABLE','I_TVACONTRE');
   MajChampFicheImmoEuro( Q,'I_TVARECUPEREE','I_TVARECUPCONTRE');
   MajChampFicheImmoEuro( Q,'I_BASETAXEPRO','I_BASETPCONTRE');
   MajChampFicheImmoEuro( Q,'I_BASEECO','I_BASEECOCONTRE');
   MajChampFicheImmoEuro( Q,'I_BASEFISC','I_BASEFISCCONTRE');
   MajChampFicheImmoEuro( Q,'I_MONTANTBASEAMORT','I_BASEAMORTCONTRE');
   MajChampFicheImmoEuro( Q,'I_VOACEDE','I_VOACEDECONTRE');
   MajChampFicheImmoEuro( Q,'I_MONTANTPREMECHE','I_PREMECHECONTRE');
   MajChampFicheImmoEuro( Q,'I_MONTANTSUIVECHE','I_SUIVECHECONTRE');
   MajChampFicheImmoEuro( Q,'I_VNC','I_VNC');
   MajChampFicheImmoEuro( Q,'I_FRAISECHE','I_FRAISECHECONTRE');
   MajChampFicheImmoEuro( Q,'I_RESIDUEL','I_RESIDUELCONTRE');
   MajChampFicheImmoEuro( Q,'I_DEPOTGARANTIE','I_DEPOTGARCONTRE');
   MajChampFicheImmoEuro( Q,'I_TVAAREVERSER','I_TVAREVERSCONTRE');
   MajChampFicheImmoEuro( Q,'I_REPRISEECO','I_REPRISEECOCONTRE');
   MajChampFicheImmoEuro( Q,'I_REPRISEFISCAL','I_REPRISEFICONTRE');
   MajChampFicheImmoEuro( Q,'I_BASEAMORDEBEXO','I_BASEDEBEXOCONTRE');
   MajChampFicheImmoEuro( Q,'I_BASEAMORFINEXO','I_BASEFINEXOCONTRE');
   MajChampFicheImmoEuro( Q,'I_REVISIONECO','I_REVECOCONTRE');
   MajChampFicheImmoEuro( Q,'I_REVISIONFISCALE','I_REVFISCCONTRE');
   MajChampFicheImmoEuro( Q,'I_MONTANTEXC','I_MONTANTEXCCONTRE');
   MajChampFicheImmoEuro( Q,'I_REPCEDECO','I_REPECOCONTRE');
   MajChampFicheImmoEuro( Q,'I_REPCEDFISC','I_REPFISCCONTRE');
   MajChampFicheImmoEuro( Q,'I_REINTEGRATION','I_REINTEGRATION');
   MajChampFicheImmoEuro( Q,'I_VALEURACHAT','I_VALEURACHAT');
   MajChampFicheImmoEuro( Q,'I_MONTANTEXCCED','I_MONTANTEXCCED');
   Q.FindField('I_DEVISE').AsString := 'EUR';  //EPZ 15/12/2000
   if (Q.FindField ('I_SAISIECONTRE').AsString = 'X') then Q.FindField ('I_SAISIECONTRE').AsString := '-'
   else Q.FindField ('I_SAISIECONTRE').AsString := 'X';
   if plan_encours <> - 1 then Q.FindField ('I_PLANACTIF').AsInteger := plan_encours;
end;

procedure MajPlanEcheEuro ( var Q: TQuery);
var QEche : TQuery;
begin
  if (Q.FindField ('I_NATUREIMMO').AsString = 'PRO') or
        (Q.FindField ('I_NATUREIMMO').AsString = 'FI') then exit;
  QEche := OpenSQL ('SELECT * FROM IMMOECHE WHERE IH_IMMO="'+Q.FindField ('I_IMMO').AsString+'"', False);
  while not QEche.Eof do
  begin
    QEche.Edit;
    QEche.FindField ('IH_MONTANT').AsFloat := ImPivotToEuro(QEche.FindField ('IH_MONTANT').AsFloat);
    QEche.FindField ('IH_FRAIS').AsFloat := ImPivotToEuro(QEche.FindField ('IH_FRAIS').AsFloat);
    QEche.Post;
    QEche.Next;
  end;
  Ferme (QEche);
end;


function MajImmoLogEuro ( var Q: TQuery; bInc : boolean) : integer;
begin
    Q.Edit;
    MajChampImmoEuro (Q,'IL_MONTANTCES');
    MajChampImmoEuro (Q,'IL_MONTANTECL');
    MajChampImmoEuro (Q,'IL_MONTANTDOT');
    MajChampImmoEuro (Q,'IL_VOCEDEE');
    MajChampImmoEuro (Q,'IL_TVAAREVERSER');
    MajChampImmoEuro (Q,'IL_MONTANTREI');
    MajChampImmoEuro (Q,'IL_REVISIONDOTECO');
    MajChampImmoEuro (Q,'IL_REVISIONREPECO');
    MajChampImmoEuro (Q,'IL_REVISIONREPFISC');
    MajChampImmoEuro (Q,'IL_REVISIONDOTFISC');
    MajChampImmoEuro (Q,'IL_PVALUE');
    MajChampImmoEuro (Q,'IL_VNC');
    MajChampImmoEuro (Q,'IL_REVISIONECO');
    MajChampImmoEuro (Q,'IL_REVISIONFISC');
    MajChampImmoEuro (Q,'IL_TVARECUPEREE');
    MajChampImmoEuro (Q,'IL_TVARECUPERABLE');
    MajChampImmoEuro (Q,'IL_REPRISEECO');
    MajChampImmoEuro (Q,'IL_REPRISEFISC');
    MajChampImmoEuro (Q,'IL_MONTANTEXC');
    MajChampImmoEuro (Q,'IL_MONTANTAVMB');
    MajChampImmoEuro (Q,'IL_BASEECOAVMB');
    MajChampImmoEuro (Q,'IL_BASEFISCAVMB');
    MajChampImmoEuro (Q,'IL_DOTCESSECO');
    MajChampImmoEuro (Q,'IL_DOTCESSFIS');
    MajChampImmoEuro (Q,'IL_CUMANTCESECO');
    MajChampImmoEuro (Q,'IL_CUMANTCESFIS');
    if bInc then
    begin
      Q.FindField ('IL_PLANACTIFAV').AsInteger := Q.FindField ('IL_PLANACTIFAV').AsInteger+1;
      Q.FindField ('IL_PLANACTIFAP').AsInteger := Q.FindField ('IL_PLANACTIFAP').AsInteger+1;
    end;
    Q.FindField ('IL_ORDRE').AsInteger := Q.FindField ('IL_ORDRE').AsInteger+1;
    Q.Post;
    result := Q.FindField ('IL_PLANACTIFAP').AsInteger-1;
end;

procedure MajPlanAmorEuro (var Q : TQuery; bInc : boolean);
begin
Q.Edit;
MajChampImmoEuro ( Q,'IA_MONTANTECO');
MajChampImmoEuro ( Q,'IA_MONTANTFISCAL');
MajChampImmoEuro ( Q,'IA_MONTANTDEROG');
MajChampImmoEuro ( Q,'IA_MTTECOEURO');
MajChampImmoEuro ( Q,'IA_MTTFISCALEURO');
MajChampImmoEuro ( Q,'IA_MTTDEROGEURO');
MajChampImmoEuro ( Q,'IA_BASEDEBEXOECO');
MajChampImmoEuro ( Q,'IA_BASEDEBEXOFISC');
MajChampImmoEuro ( Q,'IA_REINTEGRATION');
MajChampImmoEuro ( Q,'IA_QUOTEPART');
MajChampImmoEuro ( Q,'IA_CESSIONECO');
MajChampImmoEuro ( Q,'IA_CESSIONFISCAL');
MajChampImmoEuro ( Q,'IA_BASEDEBEXOCESS');
if bInc then Q.FindField ('IA_NUMEROSEQ').AsInteger := Q.FindField ('IA_NUMEROSEQ').AsInteger + 1;
Q.Post;
end;

(*procedure RecalculDernierPlanEuro (QImmo : TQuery);
var Q : TQuery; TotalEco,TotalFisc, EcartEco,EcartFisc : double;
    wCount : integer; LeSql: string ;
begin
TotalEco :=QImmo.FindField('I_REPRISEECO').AsFloat;
TotalFisc:=QImmo.FindField('I_REPRISEFISCAL').AsFloat;
wCount:=0;
LeSql:=' SELECT IA_MONTANTECO,IA_MONTANTFISCAL,IA_BASEDEBEXOECO,IA_BASEDEBEXOFISC,IA_MONTANTDEROG,IA_IMMO,IA_NUMEROSEQ,IA_DATE '
      +' FROM IMMOAMOR WHERE IA_IMMO="'+ QImmo.FindField('I_IMMO').AsString+'" '
      +' AND IA_NUMEROSEQ='+ QImmo.FindField('I_PLANACTIF').AsString+' ORDER BY IA_DATE' ;
Q := OpenSQL (LeSql,FALSE);
while not Q.Eof do
  begin
  Inc (wCount);
  TotalEco :=TotalEco +Q.FindField ('IA_MONTANTECO').AsFloat;
  TotalFisc:=TotalFisc+Q.FindField ('IA_MONTANTFISCAL').AsFloat;
  if (wCount=Q.RecordCount) then
    begin
    Q.Edit;
    //EcartEco:=QImmo.FindField ('I_BASEAMORFINEXO').AsFloat - TotalEco;
    EcartEco:=QImmo.FindField('I_BASEECO').AsFloat-TotalEco;
    Q.FindField ('IA_MONTANTECO').AsFloat   :=Q.FindField ('IA_MONTANTECO').AsFloat+EcartEco;
    Q.FindField ('IA_BASEDEBEXOECO').AsFloat:=Q.FindField ('IA_MONTANTECO').AsFloat;
    if (QImmo.FindField ('I_METHODEFISC').AsString<>'') then
      begin
      //EcartFisc := QImmo.FindField ('I_BASEAMORFINEXO').AsFloat - TotalFisc;
      EcartFisc:=QImmo.FindField ('I_BASEFISC').AsFloat-TotalFisc;
      Q.FindField ('IA_MONTANTFISCAL').AsFloat :=Q.FindField ('IA_MONTANTFISCAL').AsFloat+EcartFisc;
      Q.FindField ('IA_BASEDEBEXOFISC').AsFloat:=Q.FindField ('IA_MONTANTFISCAL').AsFloat;
      Q.FindField ('IA_MONTANTDEROG').AsFloat  :=Q.FindField ('IA_MONTANTFISCAL').AsFloat-Q.FindField ('IA_MONTANTECO').AsFloat;
      end;
    Q.Post;
    end;
  Q.Next;
  end;
Ferme (Q);
end;*)

procedure MajChampImmoEuro ( var Q : TQuery; sChamp : string);
begin
  Q.FindField (sChamp).AsFloat := ImPivotToEuro (Q.FindField (sChamp).AsFloat);
end;

end.
