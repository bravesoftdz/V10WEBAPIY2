unit siscotrt;

interface

uses SysUtils,Classes,HEnt1,ImEnt,(*IntegEcr,*)ImGenEcr, imoutgen, hmsgbox;

procedure GenereFichierMouvementSisco (ListeCompte : TList;ListeEcriture : TList; Fichier : string);
function MoisRelatif (Date : TDateTime; DebutExercice : TDateTime) : Word;
function CompareEcriture (Item1,Item2:Pointer) : integer;
function CompareCompte (Item1,Item2:Pointer) : integer;

const
  EnteteTRT = '***DEBUT***';
  FinTRT = '***FIN***';
  FormatDossierTRT = '%2.2s%5.5s%6.6s%6.6s%1.1s%2.2s%1.1s%2.2s%2.2s%06.6d%1.1s%1.1s%3.3s';
  FormatCompteTRT = '%1.1s%10.10s%25.25s%1.1s%1.1s%1.1s%14.14s%14.14s%3.3s%1.1s%1.1s%1.1s';
  FormatMoisTRT='%1.1s%2.2s';
  FormatJournalTRT='%1.1s%2.2s%20.20s%10.10s%1.1s%1.1s%1.1s%1.1s';
  FormatEcritureTRT='%1.1s%02.2d%010.10s%20.20s%10.10s%1.1s%1.1s%13.13s%13.13s%4.4s%6.6s%2.2s%8.8s%6.6s%1.1s%10.10s%6.6s%1.1s%3.3s%13.13s%6.6s%1.1s%1.1s%13.13s%8.8s%3.3s';
  NCOMPTE0 = '0000000000';
type
  TDossierTRT = class (TObject)
  private
    fCode : string; // char 2
    fDossier : string; // char 5
    fDebutExercice : string; // char 6
    fFinExercice : string; // char 6
    fTypeArchivage : string; // char 1
    fLongueurCompte : string; // char 2
    fPrecisionMontant : string; // char 1
    fVersionFichier : string; // char 2
    fVersionProgramme : string; // char 2
    fNbEcriture : integer;
    fFormatMontant : string; // char 1
    fOrigine : string; // char 1
    fMonnaie : string; // char 3
  public
    constructor Create;
    destructor Destroy;override;
    procedure Formate(var Ligne:string);
    property NbEcriture : Integer read fNbEcriture write fNbEcriture;
  end;

  TCompteTRT = class (TObject)
  private
    fCode : string; // char 1
    fNumero : string; // char 10
    fLibelle : string; // char 25
    fLettrageAuto : string; // char 1
    fSoldeProg : string; // char 1
    fTypeCompte : string; // char 1
    fDebitExPrec : double;
    fCreditExPrec : double;
    fTaux : string; // char 3
    fTva : string; // char 1
    fMonnaie : string; // char 1
    fFiller : string; // char 1
  public
    constructor Create;
    destructor Destroy;override;
    procedure Formate(var Ligne:string);
    procedure Ventile (C : TDefCompte);
  end;

  TMoisTRT = class
  private
    fCode : string; // char 1
    fMoisRelatif : string; // char 2
  public
    constructor Create (Mois : Word);
    destructor Destroy;override;
    procedure Formate(var Ligne:string);
  end;
  TJournalTRT = class
  private
    fCode : string; // char 1
    fCodeJournal : string; // char 2
    fLibelle : string; // char 20
    fCompteAuto : string; // char 10
    fSens : string; // char 1
    fTypeJournal : string; // char 1
    fQuantite : string; // char 1
    fEcheance : string; // char 1
  public
    constructor Create (CodeJournal : string);
    destructor Destroy;override;
    procedure Formate(var Ligne:string);
  end;
  TEcritureTRT = class (TObject)
  private
    fCode : string; // char 1
    fJour : integer;
    fCompte : string; // char 10
    fLibelle : string; // char 20
    fPiece : string; // char 10
    fLettrage : string; // char 1
    fExtraComptable : string; // char 1
    fDebit : double;
    fCredit : double;
    fStatistique : string; // char 4
    fDateEcheance : string; // char 6
    fReglement : string; // char 2
    fQuantite : string;
    fDateCreation : string; // char 6
    fTopee : string; // char 1
    fNoCheque : string; // char 10
    fDateRelance : string; // char 6
    fNiveauRelance : string; // char 6
    fDevise : string; // char 3
    fMontantDevise : string; // char 13
    fDateValeur : string; // char 6
    fSuiteLettrage : string; // char 1
    fDeviseOrigine : string; // char 1
    fContreValeur : string; // char 13
    fPointeur : string; // char 8
    fAutreMonnaie : string; // char 3
  public
    constructor Create;
    destructor Destroy; override;
    procedure Formate(var Ligne:string);
    procedure Ventile (E : TLigneEcriture);
  end;

implementation

uses Outils,ImpExp;

// ------------------------------------------------  Dossier

constructor TDossierTRT.Create;
var Year,Month,Day : Word;
begin
  fCode := '00';
  fDossier := '00001';
  DecodeDate(VHImmo^.Encours.Deb,Year,Month,Day);
  fDebutExercice := Format ('%02.2d%02.2d%2.2s',[Day,Month,Copy(IntToStr(Year),3,2)]);
  DecodeDate(VHImmo^.Encours.Fin,Year,Month,Day);
  fFinExercice := Format ('%02.2d%02.2d%2.2s',[Day,Month,Copy(IntToStr(Year),3,2)]);
  fTypeArchivage := 'E';
  fLongueurCompte := '10';
  fPrecisionMontant :=IntToStr(V_PGI.OkDecV);
  fVersionFichier := '10';
  fVersionProgramme := '06';
  fNbEcriture := 0;
  fFormatMontant := 'G';
  fOrigine := 'PL';
  fMonnaie := 'FRF';
end;

destructor TDossierTRT.Destroy;
begin
  inherited Destroy;
end;

procedure TDossierTRT.Formate(var Ligne:string);
begin
 Ligne := Format(FormatDossierTRT,[fCode,fDossier,fDebutExercice,
    fFinExercice,fTypeArchivage,fLongueurCompte,fPrecisionMontant,
    fVersionFichier,fVersionProgramme,fNbEcriture,fFormatMontant,
    fOrigine,fMonnaie]);
end;

// ------------------------------------------------  Compte

constructor TCompteTRT.Create;
begin
  fCode := 'C';
  fNumero := '          ';
  fLibelle := '                         ';
  fLettrageAuto := ' ';
  fSoldeProg := ' ';
  fTypeCompte :=' ';
  fDebitExPrec := 0.00;
  fCreditExPrec := 0.00;
  fTaux := '   ';
  fTva := ' ';
  fMonnaie :=' ';
  fFiller :=' ';
end;

destructor TCompteTRT.Destroy;
begin
  inherited Destroy;
end;

procedure TCompteTRT.Formate(var Ligne:string);
var sDebit, sCredit :string;
begin
 sDebit := Format('%014.14d',[StrToInt(FloatToStr(fDebitExPrec*100))]);
 sCredit := Format('%014.14d',[StrToInt(FloatToStr(fCreditExPrec*100))]);
 Ligne := Format(FormatCompteTRT,[fCode,fNumero,fLibelle,fLettrageAuto,
          fSoldeProg,fTypeCompte,sDebit,sCredit,fTaux,fTva,
          fMonnaie,fFiller]);
end;

procedure TCompteTRT.Ventile (C : TDefCompte);
begin
  fNumero := C.Compte+Copy(NCOMPTE0,0,10-Length(C.Compte));
  fLibelle := C.Libelle;
end;


// ------------------------------------------------  Mois

constructor TMoisTRT.Create (Mois : Word);
begin
  fCode := 'M';
  fMoisRelatif := Format ('%02.2d',[Mois]);
end;

destructor TMoisTRT.Destroy;
begin
  inherited Destroy;
end;

procedure TMoisTRT.Formate(var Ligne:string);
begin
 Ligne := Format(FormatMoisTRT,[fCode,fMoisRelatif]);
end;

// ------------------------------------------------  Journal

constructor TJournalTRT.Create (CodeJournal : string);
begin
  fCode := 'J';
  fCodeJournal := Copy(CodeJournal,1,2);
  fLibelle := '                    ';
  fCompteAuto := '          ';
  fSens := ' ';
  fTypeJournal := ' ';
  fQuantite := ' ';
  fEcheance := ' ';
end;

destructor TJournalTRT.Destroy;
begin
  inherited Destroy;
end;

procedure TJournalTRT.Formate(var Ligne:string);
begin
 Ligne := Format(FormatJournalTRT,[fCode,fCodeJournal,fLibelle,
       fCompteAuto,fSens,fTypeJournal,fQuantite,fEcheance]);
end;

// ------------------------------------------------  Ecriture

constructor TEcritureTRT.Create;
var Year,Month,Day : Word;
begin
  fCode := 'E';
  fJour := 0;
  fCompte := '          ';
  fLibelle := '                    ';
  fPiece := '          ';
  fLettrage := ' ';
  fExtraComptable := ' ';
  fDebit := 0.00;
  fCredit := 0.00;
  fStatistique := '    ';
  fDateEcheance := '      ';
  fReglement := '  ';
  fQuantite := '       ';
  DecodeDate(date,Year,Month,Day);
  fDateCreation := Format ('%02.2d%02.2d%2.2s',[Day,Month,Copy(IntToStr(Year),3,2)]);
  fTopee := ' ';
  fNoCheque := '          ';
  fDateRelance := '     ';
  fNiveauRelance := '      ';
  fDevise := '   ';
  fMontantDevise := '             ';
  fDateValeur := '      ';
  fSuiteLettrage := ' ';
  fDeviseOrigine := 'F';
  fContreValeur := '             ';
  fPointeur := '        ';
  fAutreMonnaie := '   ';
end;

destructor TEcritureTRT.Destroy;
begin
  inherited Destroy;
end;

procedure TEcritureTRT.Formate(var Ligne:string);
var sDebit, sCredit : string;
begin
 sDebit := Format('%013.13d',[StrToInt(FloatToStr(fDebit*100))]);
 sCredit := Format('%013.13d',[StrToInt(FloatToStr(fCredit*100))]);
 Ligne := Format(FormatEcritureTRT,[fCode,fJour,fCompte,fLibelle,fPiece,
            fLettrage,fExtraComptable,sDebit,sCredit,fStatistique,fDateEcheance,
            fReglement,fQuantite,fDateCreation,fTopee,fNoCheque,fDateRelance,
            fNiveauRelance,fDevise,fMontantDevise,fDateValeur,fSuiteLettrage,
            fDeviseOrigine,fContreValeur,fPointeur,fAutreMonnaie]);
end;

procedure TEcritureTRT.Ventile (E : TLigneEcriture);
var jj,mm,aaaa : word;
begin
  DecodeDate(E.Date,aaaa,mm,jj);
  fJour := jj;
  fLibelle := E.Libelle;
  fCompte := E.Compte+Copy(NCOMPTE0,0,10-Length(E.Compte));
  fDebit := E.Debit;
  fCredit := E.Credit;
end;

function CompareCompte (Item1,Item2:Pointer) : integer;
var Cpt1,Cpt2 : ^TDefCompte;
begin
  Cpt1 := Item1;Cpt2 := Item2;
  if Cpt1.Compte > Cpt2.Compte then Result := 1
  else if Cpt1.Compte < Cpt2.Compte then Result := -1
  else Result := 0;
end;

function CompareEcriture (Item1,Item2:Pointer) : integer;
var mr1, mr2 : integer;
begin
  result := 0;
  if Item1=nil then result := -1;
  if Item2=nil then result :=1;
  if (Item1=nil) or (Item2=nil) then exit;
  try
    mr1 := MoisRelatif(TLigneEcriture(Item1).Date,VHImmo^.Encours.Deb);
    mr2 := MoisRelatif(TLigneEcriture(Item2).Date,VHImmo^.Encours.Deb);
    if mr1 <> mr2 then Result := mr1 - mr2 else
    begin
      if TLigneEcriture(Item1).CodeJournal > TLigneEcriture(Item2).CodeJournal then Result := 1
      else if TLigneEcriture(Item1).CodeJournal < TLigneEcriture(Item2).CodeJournal then Result := -1;
    end;
  except
    result := 0;
  end;
end;

// Procedure principale de génération du fichier de mouvement

procedure GenereFichierMouvementSisco (ListeCompte : TList;ListeEcriture : TList; Fichier : string);
var   FichierIE : TextFile;
      Enreg : string;
      SDossier : TDossierTRT;
      SCompte : TCompteTRT;
      SMois : TMoisTRT;
      SJournal : TJournalTRT;
      SEcriture : TEcritureTRT;
      MoisR,MoisEnCours : word;
      JournalEnCours : string;
      i : integer;
      EnregEcr : TLigneEcriture;
      EnregCpt : ^TDefCompte;
begin
  AssignFile(FichierIE,Fichier) ;        //  OuvrirFichierIE(FichierIE,Fichier,160);
  {$I-} Rewrite(FichierIE) ;  {$I+}
  if IoResult<>0 then
  begin
    PgiBox('Impossible d''écrire dans le fichier #10'+Fichier,'Export Sisco') ;
    Exit ;
  end ;
  try
    writeln(FichierIE,EnteteTRT);
    SDossier := TDossierTRT.Create;
    SDossier.NbEcriture := ListeEcriture.Count;
    SDossier.Formate(Enreg);
    writeln(FichierIE,Enreg);
    SDossier.Free;
    MoisEnCours := 0;
    if (ListeCompte <> nil) and (ListeCompte.Count > 0 ) then
    begin
      ListeCompte.Sort(CompareCompte);
      for i := 0 to ListeCompte.Count-1 do
      begin
        EnregCpt := ListeCompte.Items[i];
        SCompte := TCompteTRT.Create;
        SCompte.Ventile (EnregCpt^);
        SCompte.Formate(Enreg);
        writeln(FichierIE,Enreg);
        SCompte.Free;
      end;
    end;
    if (ListeEcriture <> nil) and (ListeEcriture.Count > 0) then
    begin
      ListeEcriture.Sort (CompareEcriture);
      for i := 0 to ListeEcriture.Count-1 do
      begin
        EnregEcr := ListeEcriture.Items[i];
        MoisR := MoisRelatif(EnregEcr.Date,VHImmo^.Encours.Deb);
        if MoisR <> MoisEnCours then
        begin
          SMois := TMoisTRT.Create (MoisR);
          SMois.Formate (Enreg);
          writeln(FichierIE,Enreg);
          MoisEnCours := MoisR;
          JournalEnCours := '';
          SMois.Free;
        end;
        if EnregEcr.CodeJournal <> JournalEnCours then
        begin
          SJournal := TJournalTRT.Create(EnregEcr.CodeJournal);
          SJournal.Formate (Enreg);
          writeln(FichierIE,Enreg);
          JournalEnCours := EnregEcr.CodeJournal;
          writeln(FichierIE,'F001');
          SJournal.Free;
        end;
        SEcriture := TEcritureTRT.Create;
        SEcriture.Ventile (EnregEcr);
        SEcriture.Formate(Enreg);
        writeln(FichierIE,Enreg);
        SEcriture.Free;
      end;
    end;
    writeln(FichierIE,FinTRT);
  finally
    CloseFile(FichierIE);
    AfficheCompteRenduExport(Fichier);
  end ;
end;

function MoisRelatif (Date : TDateTime; DebutExercice : TDateTime) : Word;
var aEx,mEx,jEx : word;
    a,m,j : word;
begin
  DecodeDate (DebutExercice,aEx,mEx,jEx);
  DecodeDate (Date,a,m,j);
  if a >= aEx then
  begin
    if m >= mEx then Result := m - mEx + 1
    else Result := 12 - mEx + m + 1;
  end
  else Result := 0;
end;

end.
