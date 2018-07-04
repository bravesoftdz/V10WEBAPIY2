unit ImContra;

// 20/05/1999 - CA - Ajout méthode ChargeFromCode
// 20/05/1999 - CA - Ajout méthode GetResteAPayer
// 18/07/2003 - CA - Suppression champs euro
// BTY 02/06  - Anciennes éditions à ouvrir à CWAS
// BTY 03/06  - Optimisation en Web Access
// BTY 07/07  - ConvertEcheanceIntoTranche recopie à tort la dernière échéance dans la tranche en cours
// BTY 09/07  - Suite FQ 20005 MajListeEcheances pour ajouter/supprimer des échéances
// BTY 05/10/07 - FQ 21609 NB échéances fausses + Erreur SQL clé en double avec NOMBREMOIS appelé par RecalculTranches->CalculNombreEcheance
// => dans Outils CalculNombreEcheance, remplacé NOMBREMOIS par IntervalleNbMois
interface

uses classes,
     HEnt1,
     Hctrls,
     UTob
   {$IFDEF eAGLClient}
   {$ELSE}
   {$IFNDEF DBXPRESS},dbtables{$ELSE},uDbxDataSet{$ENDIF}
   {$ENDIF eAGLClient}
   // BTY 03/06
   //{$IFDEF eAGLServer}
   //,ImplanInfo
   //{$ENDIF}
;

type
  PTranche = ^ATranche;
  ATranche = record
    DateDebut, DateFin: TDateTime;
    nEcheance : integer;
    Montant,Frais : double;
  end;
  PEcheance = ^AEcheance;
  AEcheance = record
    Date: TDateTime;
    Montant,Frais:double;
  end;
  TImContrat = class
    private
      sNature : string;
      bAvance : boolean;
      nPeriode : integer;
      dtPremEch : TDateTime;
      dtDernEch : TDateTime;
//      dtDebut : TDateTime;
//      dtFin : TDateTime;
      mtPremEch : double;
      mtSuivEch : double;
      mtFraisEch : double;
      procedure DetruitListeEcheances;
      procedure DetruitListeTranches;
      // procedure ClearListeEcheances; à publier pour FQ 20005
      procedure CLearListeTranches;
      procedure RecalculTranches;
      function IsEcheance : boolean;
    public
      sCode : string;
      sTypeLoyer : string;
      ListeEcheances : TList;
      ListeTranches : TList;
      constructor Create;
      destructor Destroy;override;
      // BTY 02/06 Anciennes éditions à ouvrir à CWAS
      {$IFDEF EAGLSERVER}
       // BTY 03/06
       //procedure ChargeFromCodeWA (PlanInfo:TPlanInfo; CodeImmo:string);
       //procedure ChargeWA (PlanInfo:TPlanInfo; Q:TQuery);
      {$ELSE}
       // BTY 03/06
       //procedure Charge (Q : TQuery);
      procedure ChargeTOB (OB : TOB);
      {$ENDIF}
      procedure ClearListeEcheances;  // Publier
      procedure Charge (Q : TQuery);
      procedure ChargeFromCode (CodeImmo : string);
      procedure SetPeriode (Periodicite : String);
      procedure SetCode (Code : String);
      procedure SetNature (Nature : String);
      procedure SetTypeVersement (TypeVers : string);
      procedure AjouteTranche ( dtDeb,dtFin: TDateTime;nEch: integer;mtEch,mtFrais : double);
      procedure ConvertTrancheIntoEcheance;
      procedure MetTrancheDansEcheance;  // ajout mbo pour import expert Winner contrat
      procedure ConvertEcheanceIntoTranche ;
      procedure CalculEcheances;
      procedure MajTableEcheances;
      procedure MajTableEcheancesTOB (OBEche : TOB);
      procedure ChargeTableEcheance;
      // fq 19913 - mbo
      procedure ChargeEchSuivantDate(DateCalcul:TDateTime;Avance : boolean) ;
      procedure RazEcheances;
      function GetDateDebutEcheance : TDateTime;
      function GetDateFinEcheance : TDateTime;
      function GetDateDebutContrat : TDateTime;
      function GetDateFinContrat : TDateTime;
      function SommeDesLoyers : double;
      function GetResteAPayer (dt : TDateTime) : double;
      function GetSommeDesLoyers (dtDebut,dtFin : TDateTime) : double;
      procedure MajListeEcheances (DateFin : TDateTime; NbEcheancesAjout : Integer); // FQ 20005 Suite 09/07
    end;

function CalculDateContrat (DateEcheance : TDateTime;Periode : integer; bAvance : boolean; bDebut : boolean) : TDateTime;
{$IFDEF EAGLSERVER}
{$ELSE}
function GetCoutContrat(Q : TQuery) : double;
{$ENDIF}

implementation

uses  ImOuPlan
      , Outils
      , SysUtils  {05/10/07}
      {$IFNDEF EAGLSERVER}
      ,Planeche
      {$ENDIF}
      ;

constructor TImContrat.Create;
begin
  inherited Create;
  ListeEcheances := TList.Create;
  ListeTranches := TList.Create;
  ListeEcheances.Clear;
  ListeTranches.Clear;
end;

destructor TImContrat.Destroy;
begin
  DetruitListeEcheances;
  DetruitListeTranches;
  inherited Destroy;
end;

// BTY 03/06 Optimisation
//{$IFDEF EAGLSERVER}
//{$ELSE}
procedure TImContrat.Charge (Q : TQuery);
var
  QQ : TQuery;
  sTVADefault : string;
  dRet : double;
begin
  SetTypeVersement (Q.FindField('I_VERSEMENTCB').AsString);
  sTypeLoyer := Q.FindField('I_TYPELOYERCB').AsString;
  dtPremEch := Q.FindField('I_DATEDEBUTECH').AsDateTime;
  dtDernEch := Q.FindField('I_DATEFINECH').AsDateTime;
//EPZ 19/12/2000 : il faut stocker les échéances HT récupérables
//  mtPremEch := Q.FindField('I_MONTANTPREMECHE').AsFloat;
//  mtSuivEch := Q.FindField('I_MONTANTSUIVECHE').AsFloat;
  sTVADefault := '';
  if Q.FindField('I_TVARECUPERABLE').AsFloat=0 then
  begin
    QQ := OpenSQL ('SELECT G_TVA FROM GENERAUX WHERE (G_GENERAL="'+Q.FindField('I_COMPTELIE').AsString+'")',TRUE);
    if not QQ.EOF then sTVADefault := QQ.FindField('G_TVA').AsString;
    Ferme(QQ);
  end;
  mtPremEch := CalculEchTva(Q, sTVADefault,  Q.FindField('I_MONTANTPREMECHE').AsFloat, dRet);
  mtSuivEch := CalculEchTva(Q, sTVADefault,  Q.FindField('I_MONTANTSUIVECHE').AsFloat, dRet);

//EPZ 19/12/2000
  mtFraisEch := Q.FindField('I_FRAISECHE').AsFloat;
  SetPeriode (Q.FindField('I_PERIODICITE').AsString);
  SetNature (Q.FindField('I_NATUREIMMO').AsString);
  SetCode (Q.FindField('I_IMMO').AsString);
end;
//{$ENDIF}


// BTY 03/06
// BTY 02/06 Pour les éditions en Web Access, nelles fonctions
//{$IFDEF EAGLSERVER}
{procedure TImContrat.ChargeFromCodeWA (PlanInfo:TPlanInfo; CodeImmo:string);
var Q : TQuery;
begin
  Q := OpenSQL ('SELECT I_VERSEMENTCB,I_TYPELOYERCB,I_DATEDEBUTECH,'+
                'I_TVARECUPERABLE,I_TVARECUPEREE,I_COMPTELIE,I_ORGANISMECB,I_MONTANTHT,'+
                'I_DATEFINECH,I_MONTANTPREMECHE,I_MONTANTSUIVECHE,I_FRAISECHE,'+
                'I_PERIODICITE,I_NATUREIMMO,I_IMMO FROM IMMO WHERE I_IMMO="'+CodeImmo+'"',True);
  if not Q.Eof then
  begin
    ChargeWA (PlanInfo, Q);
  end;
  Ferme (Q);
end;

procedure TImContrat.ChargeWA (PlanInfo:TPlanInfo; Q : TQuery);
var
  QQ : TQuery;
  sTVADefault : string;
  dRet : double;
//F : TextFile;

begin
  SetTypeVersement (Q.FindField('I_VERSEMENTCB').AsString);
  sTypeLoyer := Q.FindField('I_TYPELOYERCB').AsString;
  dtPremEch := Q.FindField('I_DATEDEBUTECH').AsDateTime;
  dtDernEch := Q.FindField('I_DATEFINECH').AsDateTime;
//EPZ 19/12/2000 : il faut stocker les échéances HT récupérables
//  mtPremEch := Q.FindField('I_MONTANTPREMECHE').AsFloat;
//  mtSuivEch := Q.FindField('I_MONTANTSUIVECHE').AsFloat;
  sTVADefault := '';
  if Q.FindField('I_TVARECUPERABLE').AsFloat=0 then
  begin
    QQ := OpenSQL ('SELECT G_TVA FROM GENERAUX WHERE (G_GENERAL="'+Q.FindField('I_COMPTELIE').AsString+'")',TRUE);
    if not QQ.EOF then sTVADefault := QQ.FindField('G_TVA').AsString;
    Ferme(QQ);
  end;
// BTY***
//AssignFile (F,   'c:\CB_WA.log');
//if   FileExists ('c:\CB_WA.log') then  Append (F)
//else Rewrite (F);
//Writeln (F, TimeToStr(time) + ' I_immo= '+Q.FindField('I_IMMO').AsString +
//         ' I_TVA_Rable= '+ Q.FindField('I_TVARECUPERABLE').AsString  +
//         ' sTVADefaut= ' + sTVADefault);
//CloseFile (F);

  mtPremEch := CalculEchTva(PlanInfo, Q, sTVADefault,  Q.FindField('I_MONTANTPREMECHE').AsFloat, dRet);
  mtSuivEch := CalculEchTva(PlanInfo, Q, sTVADefault,  Q.FindField('I_MONTANTSUIVECHE').AsFloat, dRet);

  //EPZ 19/12/2000
  mtFraisEch := Q.FindField('I_FRAISECHE').AsFloat;
  SetPeriode (Q.FindField('I_PERIODICITE').AsString);
  SetNature (Q.FindField('I_NATUREIMMO').AsString);
  SetCode (Q.FindField('I_IMMO').AsString);
end; }
//{$ENDIF}


procedure TImContrat.ChargeFromCode (CodeImmo : string);
var Q : TQuery;
begin
  Q := OpenSQL ('SELECT I_VERSEMENTCB,I_TYPELOYERCB,I_DATEDEBUTECH,'+
                'I_TVARECUPERABLE,I_TVARECUPEREE,I_COMPTELIE,I_ORGANISMECB,I_MONTANTHT,'+
                'I_DATEFINECH,I_MONTANTPREMECHE,I_MONTANTSUIVECHE,I_FRAISECHE,'+
                'I_PERIODICITE,I_NATUREIMMO,I_IMMO FROM IMMO WHERE I_IMMO="'+CodeImmo+'"',True);
// BTY 03/06
//{$IFDEF EAGLSERVER}
//{$ELSE}
  if not Q.Eof then
  begin
    Charge (Q);
  end;
//{$ENDIF}
  Ferme (Q);
end;

{$IFDEF EAGLSERVER}
{$ELSE}
procedure TImContrat.ChargeTOB (OB : TOB);
var
  QQ : TQuery;
  sTVADefault : string;
  dRet : double;
begin
  SetTypeVersement (OB.GetValue('I_VERSEMENTCB'));
  sTypeLoyer := OB.GetValue('I_TYPELOYERCB');
  dtPremEch := OB.GetValue('I_DATEDEBUTECH');
  dtDernEch := OB.GetValue('I_DATEFINECH');
//EPZ 19/12/2000 : il faut stocker les échéances HT récupérables
//  mtPremEch := OB.GetValue('I_MONTANTPREMECHE');
//  mtSuivEch := OB.GetValue('I_MONTANTSUIVECHE');
  sTVADefault := '';
  if OB.GetValue('I_TVARECUPERABLE')=0 then
  begin
    QQ := OpenSQL ('SELECT G_TVA FROM GENERAUX WHERE (G_GENERAL="'+OB.GetValue('I_COMPTELIE')+'")',TRUE);
    if not QQ.EOF then sTVADefault := QQ.FindField('G_TVA').AsString;
    Ferme(QQ);
  end;
  mtPremEch := CalculEchTvaTOB(OB, sTVADefault,  OB.GetValue('I_MONTANTPREMECHE'), dRet);
  mtSuivEch := CalculEchTvaTOB(OB, sTVADefault,  OB.GetValue('I_MONTANTSUIVECHE'), dRet);
//EPZ 19/12/2000
  mtFraisEch := OB.GetValue('I_FRAISECHE');
  SetPeriode (OB.GetValue('I_PERIODICITE'));
  SetNature (OB.GetValue('I_NATUREIMMO'));
  SetCode (OB.GetValue('I_IMMO'));
end;
{$ENDIF}

procedure TImContrat.DetruitListeEcheances;
var i:integer;ARecord:PEcheance;
begin
  if ListeEcheances <> nil then
  begin
    for i := 0 to (ListeEcheances.Count - 1) do
    begin
      ARecord := ListeEcheances.Items[i];
      Dispose(ARecord);
    end;
    ListeEcheances.Free;
    ListeEcheances := nil;
  end;
end;

procedure TImContrat.DetruitListeTranches;
var i:integer;ARecord:PTranche;
begin
  if ListeTranches <> nil then
  begin
    for i := 0 to (ListeTranches.Count - 1) do
    begin
      ARecord := ListeTranches.Items[i];
      Dispose(ARecord);
    end;
    ListeTranches.Free;
    ListeTranches := nil;
  end;
end;

// FQ 20005 Suite
// Le nb d'échéances a été modifié => supprimer ou ajouter des échéances
// et si échéances ajoutées leur appliquer le même montant/frais que la dernière échéance
procedure TImContrat.MajListeEcheances (DateFin: TDateTime; NbEcheancesAjout : Integer);
var APremier,ADernier,ARecord : PEcheance;
    i , iDernier: Integer;
begin
    // récup première et dernière échéance
    APremier := ListeEcheances.Items[0];
    ADernier := ListeEcheances.Items[ListeEcheances.Count - 1];
    iDernier := ListeEcheances.Count-1;

    // ajouter des échéances
    for i:= 1 to (NbEcheancesAjout) do
    begin
       New(ARecord);
       ARecord^.Date := PlusMois (ADernier^.Date, (nPeriode * i));
        //ARecord^.Date := PlusMois(ADatePremier, (iDernier+i) * nPeriode);
       ARecord^.Montant := ADernier^.Montant;
       ARecord^.Frais := ADernier^.Frais;
       ListeEcheances.Add(ARecord);
    end;
    // ou supprimer les échéances en trop
    for i:= iDernier downto (iDernier+NbEcheancesAjout+1) do
    begin
       ListeEcheances.Delete(i);
    end;
end;

procedure TImContrat.CalculEcheances;
var dtEnCours : TDateTime;
    ARecord : PEcheance;
    nEcheance : integer;
begin
//EPZ 19/12/2000 : il faut stocker les échéances HT récupérables

//EPZ 19/12/2000
  if sTypeLoyer = 'LCO' then
  begin
    ClearListeEcheances;
    dtEnCours := dtPremEch;
    New(ARecord);
    ARecord^.Date := dtEnCours;
    ARecord^.Montant := mtPremEch;
    ARecord^.Frais := mtFraisEch;
    if ARecord^.Date <= dtDernEch then
    begin
      ListeEcheances.Add(ARecord);
      dtEnCours := PlusMois(dtEnCours,nPeriode);
      nEcheance := 1;
      while dtEnCours <= dtDernEch do
      begin
        Inc(nEcheance,1);
        New(ARecord);
        ARecord^.Date := dtEnCours;
        ARecord^.Montant := mtSuivEch;
        ARecord^.Frais := mtFraisEch;
        if ARecord^.Date <= dtDernEch then ListeEcheances.Add(ARecord)
        else Dispose (ARecord);
        dtEnCours := PlusMois(dtPremEch,nPeriode*nEcheance);
      end;
    end
    else dispose (ARecord);
  end else ConvertTrancheIntoEcheance;  // Cas du loyer variable
end;

procedure TImContrat.ConvertEcheanceIntoTranche ;
var i,CurLigne,nEcheance:integer;
    ARecordEcheance:PEcheance;
    ARecordTranche:PTranche;
    CurMontant,CurFrais : double;
begin
  ClearListeTranches;
  ARecordTranche := nil;
  if ListeEcheances.Count > 0 then
  begin
    CurLigne := 0;CurMontant := -1;
    CurFrais := -1; i := -1; nEcheance := 0;
    repeat
      i := i+1;
      ArecordEcheance := ListeEcheances.Items[i];

      if (ARecordEcheance^.Montant <> CurMontant) or (ARecordEcheance^.Frais <> CurFrais) then
      begin
        // Enregistrer la tranche en cours
        if CurLigne > 0 then
        begin
          ARecordTranche^.nEcheance := nEcheance;
          if bAvance then
             ARecordTranche^.DateFin := ARecordEcheance^.Date - 1
          else
             ARecordTranche^.DateFin := PlusMois(ARecordEcheance^.Date, (-1)*nPeriode);
          ListeTranches.Add(ARecordTranche);
        end;
        // Démarrer une nelle tranche
        nEcheance := 1;
        CurMontant := ARecordEcheance^.Montant;
        CurFrais := ARecordEcheance^.Frais;
        CurLigne := CurLigne + 1;
        New (ARecordTranche);
        if bAvance then ARecordTranche^.DateDebut := ARecordEcheance^.Date
        else ARecordTranche^.DateDebut := PlusMois(ARecordEcheance^.Date,(-1)*nPeriode)+1;
        ARecordTranche^.Montant := ARecordEcheance^.Montant;
        ARecordTranche^.Frais := ARecordEcheance^.Frais;
      end
      else
        nEcheance := nEcheance + 1;

    until (i = (ListeEcheances.Count - 1));

{      if (ARecordEcheance^.Montant <> CurMontant) or (ARecordEcheance^.Frais <> CurFrais)
          or (i = (nLigne - 1)) then
      begin
        if CurLigne > 0 then
        begin
          if (i = (nLigne - 1)) then nEcheance := nEcheance + 1;
          ARecordTranche^.nEcheance := nEcheance;
          if bAvance then
          begin
              if (i = (nLigne - 1)) then ARecordTranche^.DateFin := PlusMois(ARecordEcheance^.Date,nPeriode)-1
              else ARecordTranche^.DateFin := ARecordEcheance^.Date - 1
          end
          else
          begin
              if (i = (nLigne - 1)) then ARecordTranche^.DateFin := ARecordEcheance^.Date
              else ARecordTranche^.DateFin := PlusMois(ARecordEcheance^.Date, (-1)*nPeriode);
          end;
          ListeTranches.Add(ARecordTranche);
        end;
        if (i < (ListeEcheances.Count - 1)) then
        begin
          nEcheance := 1;
          CurMontant := ARecordEcheance^.Montant;
          CurFrais := ARecordEcheance^.Frais;
          CurLigne := CurLigne + 1;
          New (ARecordTranche);
          if bAvance then ARecordTranche^.DateDebut := ARecordEcheance^.Date
          else ARecordTranche^.DateDebut := PlusMois(ARecordEcheance^.Date,(-1)*nPeriode)+1;
          ARecordTranche^.Montant := ARecordEcheance^.Montant;
          ARecordTranche^.Frais := ARecordEcheance^.Frais;
        end;
      end
      else
          nEcheance := nEcheance + 1;
    until (i = (ListeEcheances.Count - 1)); }

    // Ne pas oublier la dernière tranche
    if ARecordTranche  <> nil then
    begin
      ARecordTranche^.nEcheance := nEcheance;
      if bAvance then
         ARecordTranche^.DateFin := PlusMois(ARecordEcheance^.Date,nPeriode)-1
      else
         ARecordTranche^.DateFin := ARecordEcheance^.Date;
      ListeTranches.Add(ARecordTranche);
    end;

//for i:= 0 to (ListeTranches.Count-1) do
//ARecordTranche := ListeTranches.Items[i];

  end;
end;

procedure TImContrat.ConvertTrancheIntoEcheance;
var i,j:integer;
    ARecordEcheance : PEcheance;
    ARecordTranche : PTranche;
    dtDeb:TDateTime;
begin
  if (ListeTranches.Count = 0) and (ListeEcheances.Count > 0) then
    ConvertEcheanceIntoTranche
  else
  begin
    RecalculTranches;
    ClearListeEcheances;
    for i := 0 to ListeTranches.Count - 1 do
    begin
      ARecordTranche := ListeTranches.Items[i];
      dtDeb := ARecordTranche^.DateDebut;
      for j:= 0 to ARecordTranche^.nEcheance - 1 do
      begin
        New(ARecordEcheance);
        if bAvance then ARecordEcheance^.Date := PlusMois(dtDeb, j*nPeriode)
        else ARecordEcheance^.Date := PlusMois(dtDeb, (j+1)*nPeriode) - 1;
        ARecordEcheance^.Montant := ARecordTranche^.Montant;
        ARecordEcheance^.Frais := ARecordTranche^.Frais;
        ListeEcheances.Add(ARecordEcheance);
      end;
    end;
  end;
end;

// ajout mbo pour import winner immo des contrats
procedure TImContrat.MetTrancheDansEcheance;
var i,j:integer;
    ARecordEcheance : PEcheance;
    ARecordTranche : PTranche;
    // dtDeb:TDateTime;
begin
    // modif mbo 29.12.06 ClearListeEcheances;
    for i := 0 to ListeTranches.Count - 1 do
    begin
      ARecordTranche := ListeTranches.Items[i];
      for j:= 0 to ARecordTranche^.nEcheance - 1 do
      begin
        New(ARecordEcheance);
        ARecordEcheance^.Date := ARecordTranche^.DateDebut;
        ARecordEcheance^.Montant := ARecordTranche^.Montant;
        ARecordEcheance^.Frais := ARecordTranche^.Frais;
        ListeEcheances.Add(ARecordEcheance);
      end;
    end;
end;


procedure TImContrat.ChargeTableEcheance ;
var
  ARecord : PEcheance;
  Q : TQuery;
begin
  ClearListeEcheances;
  Q:=OpenSQL ('SELECT IH_DATE,IH_MONTANT,IH_FRAIS FROM IMMOECHE WHERE IH_IMMO = "'+ sCode +'" ORDER BY IH_DATE',True);
  while not Q.Eof do
    begin
    new (ARecord);
    ARecord^.Date := Q.FindField('IH_DATE').AsDateTime;
    ARecord^.Montant := Q.FindField('IH_MONTANT').AsFloat;
    ARecord^.Frais := Q.FindField('IH_FRAIS').AsFloat;
    ListeEcheances.Add (ARecord);
    Q.Next;    //YCP OK
    end;
  Ferme(Q);
end;

// nvelle fonction pour édition - fq 19913 - mbo
procedure TImContrat.ChargeEchSuivantDate(DateCalcul:TDateTime;Avance : boolean) ;
var
  ARecord : PEcheance;
  Q : TQuery;
begin
  ClearListeEcheances;
  Q:=OpenSQL ('SELECT IH_DATE,IH_MONTANT,IH_FRAIS FROM IMMOECHE WHERE IH_IMMO = "'+ sCode +'" ORDER BY IH_DATE',True);
  while not Q.Eof do
    begin
    new (ARecord);
    ARecord^.Date := Q.FindField('IH_DATE').AsDateTime;

    if (Q.FindField('IH_DATE').AsDateTime > DateCalcul) and (Avance=false) then
       ARecord^.Montant := 0.0
    else
       ARecord^.Montant := Q.FindField('IH_MONTANT').AsFloat;

    ARecord^.Frais := Q.FindField('IH_FRAIS').AsFloat;
    ListeEcheances.Add (ARecord);
    Q.Next;    //YCP OK
    end;
  Ferme(Q);
end;

procedure TImContrat.MajTableEcheances;
var i:integer; ARecord : PEcheance; wMtt,wFrais: double ; LeSql: string ;
begin
if ListeEcheances.Count > 0 then
  begin
  ExecuteSQL('DELETE FROM IMMOECHE WHERE IH_IMMO="'+sCode+'"');
  for i:=0 to ListeEcheances.Count - 1 do
    begin
    ARecord := ListeEcheances.Items[i];
    wMtt:=ARecord^.Montant;
    wFrais:=ARecord^.Frais;
    LeSql:='INSERT INTO IMMOECHE (IH_IMMO,IH_NATUREIMMO,IH_DATE,IH_MONTANT,IH_FRAIS,IH_INTEGREECH,IH_INTEGREPAI) VALUES'
          +' ("'+sCode+'","'+sNature+'","'+usDateTime(ARecord^.Date)+'",'+StrfPoint(wMtt)+','+StrfPoint(wFrais)+',"-","-")' ;
    ExecuteSQL(LeSql) ;
    end;
  end ;
end;

procedure TImContrat.MajTableEcheancesTOB (OBEche : TOB);
var i:integer;
    ARecord : PEcheance;
    OB : TOB;
begin
  if ListeEcheances.Count > 0 then
  begin
    for i:=0 to ListeEcheances.Count - 1 do
    begin
      OB := TOB.Create ('IMMOECHE',OBEche,-1);
      ARecord := ListeEcheances.Items[i];
      OB.PutValue('IH_IMMO',sCode);
      OB.PutValue('IH_NATUREIMMO',sNature);
      OB.PutValue('IH_DATE',ARecord^.Date);
      OB.PutValue('IH_MONTANT',ARecord^.Montant);
      OB.PutValue('IH_FRAIS',ARecord^.Frais);
    end;
  end;
end;

procedure TImContrat.AjouteTranche (dtDeb,dtFin: TDateTime;nEch: integer;mtEch,mtFrais : double);
var ARecord : PTranche;
begin
  new (ARecord);
  ARecord^.DateDebut := dtDeb;
  ARecord^.DateFin := dtFin;
  ARecord^.nEcheance := nEch;
  ARecord^.Montant := mtEch;
  ARecord^.Frais := mtFrais;
  ListeTranches.Add (ARecord);
end;

procedure TImContrat.SetPeriode (Periodicite : String);
begin
  if Periodicite = 'MEN' then nPeriode := 1
  else if Periodicite = 'TRI' then nPeriode := 3
  else if Periodicite = 'SEM' then nPeriode := 6
  else if Periodicite = 'ANN' then nPeriode := 12
  else nPeriode := 0;
end;

procedure TImContrat.SetCode (Code : string);
begin
  sCode := Code;
end;

procedure TImContrat.SetNature (Nature : string);
begin
  sNature := Nature;
end;

procedure TImContrat.SetTypeVersement (TypeVers : string);
begin
  bAvance := (TypeVers = 'AVA');
end;

function TImContrat.GetDateDebutEcheance : TDateTime;
begin
  if (ListeEcheances.Count > 0) then
    result := PEcheance(ListeEcheances.Items[0]).Date
  else result := dtPremEch;
end;

function TImContrat.GetDateFinEcheance : TDateTime;
begin
  if (ListeEcheances.Count > 0) then
    result := PEcheance(ListeEcheances.Items[ListeEcheances.Count - 1]).Date
  else  result := dtDernEch;
end;

function TImContrat.GetDateDebutContrat : TDateTime;
begin
  if IsEcheance then
    result := CalculDateContrat(PEcheance(ListeEcheances.Items[0])^.Date,nPeriode,bAvance,true)
  else result := CalculDateContrat(dtPremEch,nPeriode,bAvance,true);
end;

function TImContrat.GetDateFinContrat : TDateTime;
begin
  if IsEcheance then
    result := CalculDateContrat(PEcheance(ListeEcheances.Items[ListeEcheances.Count - 1])^.Date,nPeriode,bAvance, false)
  else result := CalculDateContrat(dtDernEch,nPeriode,bAvance,true);
end;

function TImContrat.IsEcheance : boolean;
begin
  result := (ListeEcheances.Count > 0);
end;

procedure TImContrat.RazEcheances;
begin
  ClearListeEcheances;
  ClearListeTranches;
end;

function CalculDateContrat (DateEcheance : TDateTime;Periode : integer; bAvance : boolean; bDebut : boolean) : TDateTime;
begin
  if bAvance then  // Avance
  begin
    if bDebut then result := DateEcheance
    else result := PlusMois(DateEcheance,Periode) - 1;
  end
  else  // Echu
  begin
    if bDebut then result := PlusMois(DateEcheance,(-1)*Periode) + 1
    else result := DateEcheance;
  end;
end;

function TImContrat.SommeDesLoyers : double;
var i : integer;
   ARecord : PEcheance;
begin
  result := 0.0;
  for i:=0 to ListeEcheances.Count - 1 do
  begin
     ARecord := ListeEcheances.Items [i];
     result := result + ARecord^.Montant;
  end;
end;

procedure TImContrat.RecalculTranches;
var i:integer;
    ARecordTranche : PTranche;
begin
  for i := 0 to ListeTranches.Count - 1 do
  begin
    ARecordTranche := ListeTranches.Items[i];
    // FQ 21609 CalculNombreEcheance revu dans Outils
    ARecordTranche^.nEcheance := CalculNombreEcheance (ARecordTranche^.DateDebut,
                              ARecordTranche^.DateFin, nPeriode);
  end;
end;

// function GetResteAPayer (dt : TDateTime) : double
// Retourne le montant restant à payer après la date dt.
// Attention : ne prend pas en compte la valeur résiduelle
function TImContrat.GetResteAPayer (dt : TDateTime) : double;
var i:integer;
    ARecord : PEcheance;
    ResteAPayer : double;
begin
  ResteAPayer := 0;
  for i := 0 to ListeEcheances.Count - 1 do
  begin
    ARecord := ListeEcheances.Items[i];
    if (ARecord.Date>dt) then ResteAPayer :=  ResteAPayer + ARecord.Montant;
  end;
  Result := ResteAPayer;
end;

// GetSommeDesLoyers (dtDebut,dtFin : TDateTime) : double
// Retourne le montant des loyers entre dtDebut et deFin
// Attention : ne prend pas en compte la valeur résiduelle
function TImContrat.GetSommeDesLoyers (dtDebut,dtFin : TDateTime) : double;
var i:integer;
    ARecord : PEcheance;
    Somme : double;
begin
  Somme := 0;
  if dtDebut = 0 then
  begin
    for i := 0 to ListeEcheances.Count - 1 do
    begin
      ARecord := ListeEcheances.Items[i];
      if (ARecord.Date<=dtFin)  then
        Somme :=  Somme + ARecord.Montant
      else break;
    end;
  end else
  if dtFin = 0 then
  begin
    for i := 0 to ListeEcheances.Count - 1 do
    begin
      ARecord := ListeEcheances.Items[i];
      if (ARecord.Date>=dtDebut)  then
        Somme :=  Somme + ARecord.Montant;
    end;
  end else
  begin
    for i := 0 to ListeEcheances.Count - 1 do
    begin
      ARecord := ListeEcheances.Items[i];
      if (ARecord.Date<=dtFin) and (ARecord.Date>=dtDebut)  then
        Somme :=  Somme + ARecord.Montant;
    end;
  end;
  Result := Somme;
end;

//EPZ 04/12/00
{$IFDEF EAGLSERVER}
{$ELSE}
function GetCoutContrat(Q : TQuery) : double;
var fContrat : TImContrat;
begin
  fContrat := TImContrat.Create;
  fContrat.Charge (Q);
  fContrat.ChargeTableEcheance;
  fContrat.CalculEcheances;
  result := fContrat.SommeDesLoyers + Q.FindField('I_RESIDUEL').AsFloat;
  fContrat.free;
end;
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 20/08/2004
Modifié le ... :   /  /
Description .. : Efface la liste des échéances
Mots clefs ... :
*****************************************************************}
procedure TImContrat.ClearListeEcheances;
var
  i:integer;
  ARecord:PEcheance;
begin
  if ListeEcheances <> nil then
  begin
    for i := 0 to (ListeEcheances.Count - 1) do
    begin
      ARecord := ListeEcheances.Items[i];
      Dispose(ARecord);
    end;
    ListeEcheances.Clear;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 20/08/2004
Modifié le ... :   /  /
Description .. : Efface la liste des tranches
Mots clefs ... :
*****************************************************************}
procedure TImContrat.ClearListeTranches;
var
  i:integer;
  ARecord:PTranche;
begin
  if ListeTranches <> nil then
  begin
    for i := 0 to (ListeTranches.Count - 1) do
    begin
      ARecord := ListeTranches.Items[i];
      Dispose(ARecord);
    end;
  end;
  ListeTranches.Clear;
end;

end.
