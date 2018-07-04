unit AMIMMO;

// BTY - FQ 17259 - 01/06 Nouveau top dépréciation dans IMMO
// TGA - 26/01/2006 Ajout champs sur table IMMO
// TGA - 04/04/2006 Ajout champs sur table IMMO
// BTY - FQ 17516 - 04/06 Nouveau top changement de regroupement dans IMMO
// TGA - 07/04/06 Initialisation de champs IMMO, IMMOLOG
// BTY - FQ 17629 - 04/06 Paramètre supplémentaire de cession = règle de calcul de la PMValue
// TGA - 28/06/2006 suppression des enregs de immomvtd
// TGA - 07/09/2006 I_PRF => I_PFR
// MBO - 23/03/2007 - fq 17512 - ajout des nouveaux champs V8
// BTY - 05/07 FQ 20256 En réinitialisation d'immo, mettre le type de dérogatoire
// MBO - FQ 21754 - 30/10/07 - Modif pour nouveau paramètre amort du jour de cession
interface

uses
    SysUtils,
    {$IFDEF EAGLCLIENT}
    {$ELSE}
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
    {$ENDIF}
    uTOB,
    HEnt1,
    HCtrls,
    Math,
{$IFNDEF CMPGIS35}
    AMSYNTHESEDPI_TOF,
{$ENDIF}
    HMsgBox ;

type
  TAMImmo = class
    private
      fFiche  : TOB;
      fPlan   : TOB;
      fLog    : TOB;
      function GetMontantHT : double;
    public
      constructor Create ;
      destructor  Destroy; override;
      function    Select ( stCode : string ) : boolean;
      function    Enregistre ( bRaz : boolean = False ) : boolean;
      procedure   ChangeCode ( stNouveauCode : string );
      procedure   Reinitialise;
    published
      property MontantHT : double read GetMontantHT;
  end;

function ReinitImmo ( stCode : string ) : boolean;

implementation

{ uses Amortissement }

uses Outils,
     ImEnt,
     ImPlan,
     ImPlanInfo,
     AmSortie;

{ Fin uses Amortissement }

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 09/10/2003
Modifié le ... :   /  /
Description .. : Réinitialisation d'une immobilisation : cette opération
Suite ........ : consiste à recréer une immobilisation sans historique ni
Suite ........ : opération pour repartir d'une situation saine.
Mots clefs ... :
*****************************************************************}
function ReinitImmo ( stCode : string ) : boolean;
var UneImmo : TAMImmo;
    LaSortie : TAmSortie;
    psTVA, psPrix, psMontantExc : double;
    psDateOp : TDateTime;
    psMotif, psModeCalc, psTypeExc : string;
    bSortie : boolean;
    psRegleCession : string;
    psAmortJsortie : string;  //ajout mbo - fq 21754
begin
  Result := True; bSortie := False; psMontantExc := 0; psTVA := 0; psPrix := 0;psDateOp := iDate1900;
  { Chargement de l'immobilisation }
  UneImmo := TAmImmo.Create;
  try
    if not UneImmo.Select ( stCode ) then
    begin
      UneImmo.Free;
      exit;
    end;
    { Récupération des paramètres éventuels de sortie }
    LaSortie := TAmSortie.Create (stCode);
    try
      if LaSortie.EstSortie then
      begin
        bSortie := True;
        psTVA := LaSortie.TVA;
        psDateOp :=LaSortie.DateOp;
        psMotif := LaSortie.Motif;
        psModeCalc := LaSortie.ModeCalc;
        psPrix := LaSortie.Prix;
        psMontantExc := LaSortie.MontantExc;
        psTypeExc := LaSortie.TypeExc;
        psRegleCession := LaSortie.RegleCession;   // BTY 04/06 FQ 17629
        psAmortJsortie := LaSortie.AmortJsortie;  // mbo 30/10/2007 - fq 21754
      end;
    finally
      FreeAndNil(LaSortie);
    end;
    { Recalcul des informations }
    UneImmo.Reinitialise;
    { Enregistrement de l'immobilisation réinitialisée }
    UneImmo.Enregistre ( True );
    { Génération de la sortie }
    if bSortie then
    begin
      LaSortie := TAmSortie.Create (stCode);
      // BTY 04/06 FQ 17629 Paramètre supplémentaire = règle de calcul de la PMValue
      //LaSortie.Init ( psDateOp, psMotif, psModeCalc, psPrix, PsTVA);
      //mbo - 30/10/07 - fq 21754  ajout du paramètre psAmortJsortie
      LaSortie.Init ( psDateOp, psMotif, psModeCalc, psPrix, PsTVA, psRegleCession, psAmortJsortie);
      if psMontantExc > 0  then
      begin
        psTypeExc := 'DOT';
      end else if psMontantExc < 0 then
      begin
        psTypeExc := 'RDO';
        psMontantExc := (-1)*psMontantExc;
      end;
      LaSortie.InitExceptionnel ( psMontantExc, psTypeExc  );
      LaSortie.Execute;
      FreeAndNil(LaSortie);
    end;
    Result := (UneImmo.MontantHT <> 0) or bSortie;
  finally
    UneImmo.Free;
  end;
end;

{ TIMMO }

procedure TAMImmo.ChangeCode(stNouveauCode: string);
var i : integer;
begin
  fFiche.Detail[0].PutValue('I_IMMO', stNouveauCode);
  for i := 0 to fPlan.Detail.Count - 1 do
    fPlan.Detail[i].PutValue('IA_IMMO',stNouveauCode);
  for i := 0 to fLog.Detail.Count - 1 do
    fLog.Detail[i].PutValue('IL_IMMO',stNouveauCode);
end;

constructor TAMImmo.Create;
begin
  fFiche  := TOB.Create('', nil, -1);
  fPlan   := TOB.Create('', nil, -1);
  fLog    := TOB.Create('', nil, -1);
end;

destructor TAMImmo.Destroy;
begin
  fFiche.Free;
  fPlan.Free;
  fLog.Free;
  inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 10/10/2003
Modifié le ... :   /  /
Description .. : Enregistrement de l'immobilisation en base
Mots clefs ... :
*****************************************************************}
function TAMImmo.Enregistre ( bRaz : boolean = False ) : boolean;
begin
  BeginTrans;
  try
    if bRAz then
    begin
      { Destruction physique de l'immobilisation }
      ExecuteSQL ('DELETE FROM IMMO WHERE I_IMMO="'+fFiche.Detail[0].GetValue('I_IMMO')+'"');

      //Tga 28/06/2006 Maj Immomvtd suppression de l'immobilisation
{$IFNDEF CMPGIS35}
      AM_MAJ_IMMOMVTD('S',fFiche.Detail[0].GetValue('I_IMMO'),'',0);
{$ENDIF}

      ExecuteSQL ('DELETE FROM IMMOAMOR WHERE IA_IMMO="'+fFiche.Detail[0].GetValue('I_IMMO')+'"');
      ExecuteSQL ('DELETE FROM IMMOLOG WHERE IL_IMMO="'+fFiche.Detail[0].GetValue('I_IMMO')+'"');
    end;
    if bRaz or (not ExisteSQL ('SELECT I_IMMO FROM IMMO WHERE I_IMMO="'+FFiche.Detail[0].GetValue('I_IMMO')+'"')) then
    begin
      FFiche.InsertDB(nil);
      FPlan.InsertDB(nil);
      FLog.InsertDB(nil);
    end else
    begin
      FFiche.UpdateDB;
      FPlan.UpdateDB;
      FLog.UpdateDB;
    end;
    CommitTrans;
  except
    V_PGI.IoError:=oeUnknown;
    Rollback;
  end;
  Result := (V_PGI.IoError<>oeUnknown);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 10/10/2003
Modifié le ... :   /  /
Description .. : Réinitialisation de l'immobilisation
Mots clefs ... :
*****************************************************************}
function TAMImmo.GetMontantHT: double;
begin
  Result := fFiche.Detail[0].GetValue('I_MONTANTHT');
end;

procedure TAMImmo.Reinitialise;
var T           : TOB;
    Plan        : TPlanAmort;
    PlanInfo    : TPlanInfo;
    m,a,NbMois  : Word ;
    Q           : TQuery;
begin
  { Réinitialisation de la fiche }
  fFiche.Detail[0].PutValue('I_OPERATION','-');
  fFiche.Detail[0].PutValue('I_OPEMUTATION','-');
  fFiche.Detail[0].PutValue('I_OPEECLATEMENT','-');
  fFiche.Detail[0].PutValue('I_OPECESSION','-');
  fFiche.Detail[0].PutValue('I_OPECHANGEPLAN','-');
  fFiche.Detail[0].PutValue('I_OPELIEUGEO','-');
  fFiche.Detail[0].PutValue('I_OPEETABLISSEMENT','-');
  fFiche.Detail[0].PutValue('I_OPELEVEEOPTION','-');
  fFiche.Detail[0].PutValue('I_OPEMODIFBASES','-');
  // BTY FQ 17259 01/06 Nouveau top dépréciation
  fFiche.Detail[0].PutValue('I_OPEDEPREC','-');
  // BTY FQ 17516 04/06 Nouveau top changement de regroupement
  fFiche.Detail[0].PutValue('I_OPEREG','-');
  PlanInfo := TPlanInfo.Create('');
  PlanInfo.ChargeImmo(fFiche.Detail[0].GetValue('I_IMMO'));
  PlanInfo.Calcul(VHImmo^.Encours.Fin,true, false);
  fFiche.Detail[0].PutValue('I_MONTANTHT',PlanInfo.GetValeurAchat(VHImmo^.EnCours.Deb, False));
  fFiche.Detail[0].PutValue('I_BASEECO',PlanInfo.BaseEco);
  if fFiche.Detail[0].GetValue('I_METHODEFISC')<> '' then
    fFiche.Detail[0].PutValue('I_BASEFISC',PlanInfo.BaseEco);
  fFiche.Detail[0].PutValue('I_VNC',0);
  fFiche.Detail[0].PutValue('I_VOACEDE',0);
  fFiche.Detail[0].PutValue('I_DATECESSION',iDate1900);
  fFiche.Detail[0].PutValue('I_QTCEDE',0);
  fFiche.Detail[0].PutValue('I_IMMOORIGINEECL','');
  // TGA 26/01/2006
  fFiche.Detail[0].PutValue('I_REPRISEDEP',0);
  fFiche.Detail[0].PutValue('I_ECCLEECR','');
  fFiche.Detail[0].PutValue('I_DOCGUID','');
  fFiche.Detail[0].PutValue('I_DATEFINCB',iDate1900);
  // TGA 04/04/2006
  fFiche.Detail[0].PutValue('I_REPRISEDEPCEDEE',0);
  // TGA 10/04/2006
  fFiche.Detail[0].PutValue('I_SUSDEF','A');
  fFiche.Detail[0].PutValue('I_REGLECESSION','NOR');
  // mbo fq 17512 - fFiche.Detail[0].PutValue('I_NONDED','-');
  fFiche.Detail[0].PutValue('I_REPRISEINT',0);
  fFiche.Detail[0].PutValue('I_REPRISEINTCEDEE',0);
  fFiche.Detail[0].PutValue('I_DPI','-');
  fFiche.Detail[0].PutValue('I_DPIEC','-');
  fFiche.Detail[0].PutValue('I_CORRECTIONVR',0);
  fFiche.Detail[0].PutValue('I_CORVRCEDDE',0);
  fFiche.Detail[0].PutValue('I_SUBVENTION','NON');
  fFiche.Detail[0].PutValue('I_SBVPRI',0);
  fFiche.Detail[0].PutValue('I_SBVMTC',0);
  fFiche.Detail[0].PutValue('I_SBVPRIC',0);
  fFiche.Detail[0].PutValue('I_SBVEC','C');
  fFiche.Detail[0].PutValue('I_SBVDATE',iDate1900);
  fFiche.Detail[0].PutValue('I_CPTSBVR','');
  fFiche.Detail[0].PutValue('I_CPTSBVB','');
  // TGA 07/09/2006 fFiche.Detail[0].PutValue('I_PRF','-');
  fFiche.Detail[0].PutValue('I_PFR','-');
  fFiche.Detail[0].PutValue('I_COEFDEG',0);
  fFiche.Detail[0].PutValue('I_AMTFOR',0);
  fFiche.Detail[0].PutValue('I_AMTFORC',0);
  fFiche.Detail[0].PutValue('I_ACHFOR',iDate1900);
  fFiche.Detail[0].PutValue('I_PRIXACFORC',0);
  fFiche.Detail[0].PutValue('I_VNCFOR',0);
  fFiche.Detail[0].PutValue('I_DURRESTFOR',0);

//  fFiche.Detail[0].PutValue('I_DATEDEBECO',fFiche.Detail[0].GetValue('I_DATEAMORT'));
//  fFiche.Detail[0].PutValue('I_DATEDEBFIS',fFiche.Detail[0].GetValue('I_DATEAMORT'));

  If fFiche.Detail[0].GetValue('I_METHODEECO')='DEG' Then
    Begin
     fFiche.Detail[0].PutValue('I_DATEDEBECO',fFiche.Detail[0].GetValue('I_DATEPIECEA'));
     fFiche.Detail[0].PutValue('I_DATEDEBFIS',fFiche.Detail[0].GetValue('I_DATEPIECEA'));
    End
  Else
    Begin
     fFiche.Detail[0].PutValue('I_DATEDEBECO',fFiche.Detail[0].GetValue('I_DATEAMORT'));
     fFiche.Detail[0].PutValue('I_DATEDEBFIS',fFiche.Detail[0].GetValue('I_DATEAMORT'));
    End;

  If fFiche.Detail[0].GetValue('I_METHODEFISC')='DEG' Then
    fFiche.Detail[0].PutValue('I_DATEDEBFIS',fFiche.Detail[0].GetValue('I_DATEPIECEA'))
  Else
    fFiche.Detail[0].PutValue('I_DATEDEBFIS',fFiche.Detail[0].GetValue('I_DATEAMORT'));


  { FQ 15470 - CA - 18/03/2005 - recalcul de la reprise }
  Plan := TPlanAmort.Create(true);
  try
    Q := OpenSQL ('SELECT * FROM IMMO WHERE I_IMMO="'+fFiche.Detail[0].GetValue('I_IMMO')+'"',True);
    Plan.Charge(Q);
    Plan.Suramort := Q.FindField('I_SURAMORT').AsString='X';
    Plan.GestionFiscale := Q.FindField('I_NONDED').AsString='X';  // fq 17512 mbo
    Plan.CalculReprises;
    fFiche.Detail[0].PutValue('I_REPRISEECO', Plan.AmortEco.Reprise);
    fFiche.Detail[0].PutValue('I_REPRISEFISCAL', Plan.AmortFisc.Reprise);
    fFiche.Detail[0].PutValue('I_REPRISEDR', Plan.AmortDerog.Reprise);  // fq 17512 mbo
    fFiche.Detail[0].PutValue('I_REPRISEFEC', Plan.AmortReint.Reprise);  // fq 17512 mbo
    fFiche.Detail[0].PutValue('I_DATEDERMVTECO', Plan.GetDateFinAmortEx(Plan.AmortEco));
    fFiche.Detail[0].PutValue('I_DATEDERNMVTFISC', Plan.GetDateFinAmortEx(Plan.AmortFisc));
    Ferme (Q);
  finally
    Plan.free;
  end;
//  fFiche.Detail[0].PutValue('I_REPRISEECO',PlanInfo.GetCumulAntEco(VHImmo^.Encours.Fin));
  NOMBREMOIS (fFiche.Detail[0].GetValue('I_DATEAMORT'),VHImmo^.Encours.Deb,m,a,NbMois);
  fFiche.Detail[0].PutValue('I_DUREEREPRISE',IntToStr(MaxIntValue([0, NbMois - 1])));
//  fFiche.Detail[0].PutValue('I_REPRISEFISCAL',PlanInfo.GetCumulAntFisc(VHImmo^.Encours.Fin));
  fFiche.Detail[0].PutValue('I_REPCEDECO',0);
  fFiche.Detail[0].PutValue('I_REPCEDFISC',0);
  fFiche.Detail[0].PutValue('I_REPRISEFDRCEDEE',0); // FQ 17512 mbo
  fFiche.Detail[0].PutValue('I_REPRISEFECCEDEE',0); // FQ 17512 mbo
  fFiche.Detail[0].PutValue('I_CHANGECODE','');
  fFiche.Detail[0].PutValue('I_ETAT','OUV');
  fFiche.Detail[0].PutValue('I_TYPEEXC','');
  fFiche.Detail[0].PutValue('I_MONTANTEXC',0);
  fFiche.Detail[0].PutValue('I_PLANACTIF',0);
  if fFiche.Detail[0].GetValue('I_QUANTITE')=0 then
    fFiche.Detail[0].PutValue('I_QUANTITE',1);

  fFiche.Detail[0].PutValue('I_TYPEDEROGLIA', TypeDerogatoire(fFiche.Detail[0], nil) ); // FQ 20256

  PlanInfo.Free;
  { Suppression du plan }
  fPlan.ClearDetail;
  { Création du plan d'amortissement }
  Plan:=TPlanAmort.Create(true);
  Plan.CalculTOB( fFiche.Detail[0], iDate1900);
  Plan.SauveTOB ( fPlan );
  fFiche.Detail[0].PutValue('I_PLANACTIF',Plan.NumSeq);
  FreeAndNil( Plan );
  { Suppression de l'historique }
  fLog.ClearDetail;

  { Création de l'enregistrement d'acquisition dans le log }
  T := TOB.Create ('IMMOLOG',fLog,-1);

   // TGA 07.04.2006
  T.PutValue('IL_TAUX',0);

  T.PutValue('IL_TYPEOP','ACQ' );
  T.PutValue('IL_IMMO',fFiche.Detail[0].GetValue('I_IMMO'));
  T.Putvalue('IL_DATEOP',fFiche.Detail[0].GetValue('I_DATEPIECEA'));
  T.Putvalue('IL_ORDRE',1);
  T.Putvalue('IL_ORDRESERIE',1);
  T.Putvalue('IL_PLANACTIFAV',0);
  T.Putvalue('IL_PLANACTIFAP',1);
  T.Putvalue('IL_LIBELLE',RechDom('TIOPEAMOR', 'ACQ', FALSE)+' '+DateToStr(fFiche.Detail[0].GetValue('I_DATEPIECEA')));
  T.Putvalue('IL_TYPEMODIF',AffecteCommentaireOperation('ACQ'));
end;

function TAMImmo.Select(stCode : string): boolean;
begin
  fFiche.LoadDetailDB('IMMO', '"'+stCode+'"', '', nil, False );
  fPlan.LoadDetailDB('IMMOAMOR', '"'+stCode+'"', '', nil, False );
  fLog.LoadDetailDB('IMMOLOG', '"'+stCode+'"', '', nil, False );
  Result := ( fFiche.Detail.Count = 1 );
end;

end.
