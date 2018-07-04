{***********UNITE*************************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 10/10/2001
Modifié le ... : 10/10/2001
Description .. : But : Import/Export des balances de situation
Suite ........ : -----------------------------------------------------------------------
Suite ........ : Fonctions
Suite ........ :  - Import des balances de situation Sisco II
Suite ........ : -----------------------------------------------------------------------
Suite ........ : Utilisation :
Suite ........ :  _Importation d'une balance SiscoII
Suite ........ :        ImportBDS := TImpExpBds.Create (miSiscoII);
Suite ........ :        ImportBDS.OnInformation := OnInfoBds;
Suite ........ :        ImportBDS.Importation( FileName, CodeBal );
Suite ........ :        ImportBDS.Free;
Suite ........ : -----------------------------------------------------------------------
Mots clefs ... : BALANCE;SITUATION;SISCO;IMPORT;EXPORT
*****************************************************************}
unit ImpExpBds;

interface

uses classes, SysUtils, Math,
     {$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
     {$ENDIF}
     hctrls, uTOB,
     BalSit, LibGeneral, LibTiers, HStatus;

const
  ERR_GENE_INEXISTANT   = 1;
  ERR_TRANSACTION       = 2;
  ERR_CODEEXISTE        = 3;
  ERR_NOBALANCE         = 4;
  ERR_AUXI_INEXISTANT   = 5;
  ERR_DATE              = 6;
  ERR_PARAMSOC          = 7;

  ERROR_MSG : array[1..7] of string = (
                'Compte général inexistant.',
                'Mise à jour de la base de données impossible.',
                'Le code de la balance existe déjà.',
                'Aucune balance à importer.',
                'Compte auxiliaire inexistant.',
                '',
                'Les comptes collectifs par défaut ne sont pas renseignés.'
  );



type  TModeImpExpBds = (miSiscoII, miPGI, miHistobal);
      TImpExpBDSInfo = procedure ( Sender : TObject; Msg : string ; bErr : boolean) of object;
      TImpExpBds = class
        private
          FModeImpExp : TModeImpExpBds;
          FFileName : string;
          FCodeBal : string;
          FLibelleBal : string;
          FAbregeBal : string;
          FListeBalance : TList;
          FOnInformation : TImpExpBDSInfo;
          FLastError : integer;
          FTGen : TOB;
          FTAux : TOB;
          FDebCreColl : boolean;
          function  GetBalance ( CodeBal : string ) : TBalSit;
          function  GetExerciceFromDate (Date1, Date2 : TDateTime ) : string;
          procedure ImportSiscoII;
          procedure ImportPGI;
          procedure ImportHistobal;
          procedure InsertDatabase;
          function  CreationBalanceSiscoII ( St : string ) : boolean;
          function  RecupereCompteSiscoII ( St : string; var Gene, Auxi : string; Creation : boolean; var bCollectif : boolean) : boolean ;
          function  AjouteLigneSiscoII ( St , Gene, Auxi : string; bCollectif : boolean ) : boolean;
          function  ImportBalanceHistobal ( Exercice : string; DateDeb, DateFin : TDateTime; bTous : boolean) : boolean;
        public
          constructor Create (ModeImpExp : TModeImpExpBds);
          destructor Destroy; override;
          procedure Importation ( FileName : string; CodeBal, LibelleBal, AbregeBal : string; bSoldeDebCreColl : boolean);
          procedure Exportation ( FileName : string);
          function GetLastError : integer;
          function LastErrorMsg : string;
        published
          property OnInformation : TImpExpBDSInfo read FOnInformation write FOnInformation;
      end;

implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  CPTypeCons,
  {$ENDIF MODENT1}
  ParamSoc,
  Hent1,
  Ent1;




{ TImpExpBds }

function TImpExpBds.AjouteLigneSiscoII(St, Gene, Auxi: string; bCollectif : boolean): boolean;
var BalSit : TBalSit;
    DebitEuro, CreditEuro, Debit, Credit: double;
begin
  if Auxi <> '' then
    BalSit := GetBalance ( FCodeBal + Copy(St,12,4)+'T' )
  else BalSit := GetBalance ( FCodeBal + Copy(St,12,4)+'G' );
  if (BalSit <> nil) and (Gene <> '' )  then
  begin
    DebitEuro := Valeur(Trim(Copy(St,99,15)))/100;
    CreditEuro := Valeur(Trim(Copy(St,114,15)))/100;
    Debit   := DebitEuro;
    Credit  := CreditEuro;
    if (FDebCreColl) then
    begin
      // On ne touche à rien
      // Si le fichier Sisco est en déb. créd, on récupèrera du déb./créd.
    end else
    begin
      if Debit >= Credit then
      begin
        Debit := Debit-Credit;
        Credit := 0;
      end else
      begin
        Credit := Credit-Debit;
        Debit := 0;
      end;
    end;
    BalSit.AjouteEcriture ( Gene, Auxi, Debit, Credit );
  end;
  Result := True;
end;

constructor TImpExpBds.Create (ModeImpExp : TModeImpExpBds);
begin
  FModeImpExp := ModeImpExp;
  FListeBalance := TList.Create ;
  FTGen := TOB.Create ('', nil, - 1);
  FTAux := TOB.Create ('', nil, - 1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 12/11/2001
Modifié le ... :   /  /
Description .. : Création des balances importées au format Sisco II.
Suite ........ : Si la balance possède des auxiliaires, on crée alors dans 
Suite ........ : PGI 2 balances :
Suite ........ :   - 1 balance générale
Suite ........ :   - 1 balance auxiliaire
Mots clefs ... : BALANCE;SITUATION;CREATION
*****************************************************************}
function TImpExpBds.CreationBalanceSiscoII(St: string) : boolean;
var DateDeb, DateFin : TDateTime;
    Exercice : string;
    BalSit : TBalSit;
begin
  Result := True;
  DateDeb := Str6ToDate(Copy (St,51,6),80);
  DateFin := Str6ToDate(Copy (St,57,6),80);
  if (DateDeb=0) or (DateFin=0) then
  begin
    Result := True;
    exit;
  end;
  Exercice := GetExerciceFromDate (DateDeb,DateFin);
  if Exercice = '' then Exercice := '...';
  // On crée systématiquement une balance typée 'GEN'
  BalSit := TBalSit.Create ( FCodeBal + Copy (St,3,4)+'G' ) ;
  if BalSit.Existe then
  begin
    if Assigned(FOnInformation) then
      FOnInformation(self, TraduireMemoire('Le code de la balance existe déjà.' ), TRUE);
    fLastError := ERR_CODEEXISTE;
    BalSit.Free;
    Result := False;
    exit;
  end;
  BalSit.SetCumul('GEN');
  if FLibelleBal = '' then
    BalSit.SetLibelle ( FormatDateTime('mmmm yyyy',DateFin) )
  else BalSit.SetLibelle ( FLibelleBal + ' ('+ FormatDateTime('mm/yy',DateFin)+')');
  BalSit.SetAbrege ( FAbregeBal ) ;
  BalSit.SetExercice ( Exercice );
  BalSit.SetDateDebut( DateDeb );
  BalSit.SetDateFin ( DateFin );
  BalSit.SetDevise ( V_PGI.DevisePivot );
  // Ajout de la balance dans la liste
  FListeBalance.Add ( BalSit );

  if St[81]='O' then // les auxiliaires sont présents, on crée 1 deuxième balance
  begin
    BalSit := TBalSit.Create ( FCodeBal + Copy (St,3,4)+'T' ) ;
    if BalSit.Existe then
    begin
      if Assigned(FOnInformation) then
        FOnInformation(self, TraduireMemoire('Le code de la balance existe déjà.' ), TRUE);
      fLastError := ERR_CODEEXISTE;
      Result := False;
      exit;
    end;
    BalSit.SetCumul('G/T');
    if FLibelleBal = '' then
      BalSit.SetLibelle ( FormatDateTime('mmmm yyyy',DateFin) )
    else BalSit.SetLibelle ( FLibelleBal + ' ('+ FormatDateTime('mm/yy',DateFin)+')');
    BalSit.SetAbrege ( FAbregeBal ) ;
    BalSit.SetExercice ( Exercice );
    BalSit.SetDateDebut( DateDeb );
    BalSit.SetDateFin ( DateFin );
    BalSit.SetDevise ( V_PGI.DevisePivot );
    // Ajout de la balance dans la liste
    FListeBalance.Add ( BalSit );
  end;
end;

destructor TImpExpBds.Destroy;
var i : integer;
begin
  FTGen.Free;
  FTAux.Free;
  for i := 0 to FListeBalance.Count - 1 do
    TBalSit ( FListeBalance.Items[i] ).Free;
  inherited;
end;

procedure TImpExpBds.Exportation( FileName : string );
begin
  // Exportation balances de situation
end;

function TImpExpBds.GetBalance ( CodeBal : string ) : TBalSit;
var i : integer;
  BalSit : TBalSit;
begin
  BalSit := nil;
  for i:=0 to FListeBalance.Count - 1 do
  begin
    BalSit := TBalSit ( FListeBalance.Items[i] ) ;
    if BalSit.CodeBal = CodeBal then break
    else BalSit := nil;
  end;
  Result := BalSit;
end;

function TImpExpBds.GetExerciceFromDate(Date1, Date2 : TDateTime): string;
var Q : TQuery;
begin
  Result := '';
  Q := OpenSQL ('SELECT * FROM EXERCICE', True);
  try
    while not Q.Eof do
    begin
      if ((Q.FindField('EX_DATEDEBUT').AsDateTime <= Date1) and
                (Q.FindField('EX_DATEFIN').AsDateTime >= Date2)) then
      begin
        Result := Q.FindField('EX_EXERCICE').AsString;
        break;
      end;
      Q.Next;
    end;
  finally
    Ferme (Q);
  end;
end;

function TImpExpBds.GetLastError: integer;
begin
  Result := FLastError;
end;

procedure TImpExpBds.Importation ( FileName : string; CodeBal, LibelleBal, AbregeBal : string; bSoldeDebCreColl : boolean);
begin
  if Assigned(FOnInformation) then
    FOnInformation(self, TraduireMemoire('Importation en cours ...'), False);
  FFileName := FileName;
  FCodeBal := CodeBal;
  FLibelleBal := LibelleBal;
  FAbregeBal := AbregeBal;
  FDebCreColl := bSoldeDebCreColl;
  case FModeImpExp of
    miSiscoII : ImportSiscoII;
    miPGI : ImportPGI;
    miHistobal : ImportHistobal;
  end;
  if fLastError <> 0 then exit;
  if Transactions (InsertDatabase,3) <> oeOK then FLastError := ERR_TRANSACTION;
  if (FLastError=0) and (Assigned(FOnInformation)) then
    FOnInformation(self, TraduireMemoire('Importation terminée.'), False);
end;

function TImpExpBds.ImportBalanceHistobal(Exercice: string; DateDeb,
  DateFin: TDateTime; bTous : boolean) : boolean;
var Q : TQuery;
    BalSit : TBalSit;            
    Debit, Credit, DebitDev, CreditDev : double;
begin
  Result := True;
  Q := OpenSQL ('SELECT *,G_COLLECTIF FROM HISTOBAL LEFT JOIN GENERAUX ON G_GENERAL=HB_COMPTE1 WHERE HB_TYPEBAL="BDS" AND HB_EXERCICE="'+
                  Exercice+'" AND HB_DATE1="'+USDateTime(DateDeb)+'" AND HB_DATE2="'+
                  USDateTime(DateFin)+'"',True);
  if not Q.Eof then
  begin
    { Création de l'entête de la balance }
    if not bTous then BalSit := TBalSit.Create ( FCodeBal )
    else BalSit := TBalSit.Create ( FCodeBal+FormatDateTime('mmyy',DateFin ));
    if BalSit.Existe then
    begin
      if Assigned(FOnInformation) then
        FOnInformation(self, TraduireMemoire('Le code de la balance existe déjà.' ), TRUE);
      fLastError := ERR_CODEEXISTE;
      BalSit.Free;
      Ferme (Q);
      Result := False;
      exit;
    end;
    BalSit.SetCumul('GEN');
    if FLibelleBal = '' then
      BalSit.SetLibelle ( FormatDateTime('mmmm yyyy',DateFin) )
    else if bTous then BalSit.SetLibelle ( FLibelleBal + ' ('+ FormatDateTime('mm/yy',DateFin)+')')
      else BalSit.SetLibelle(FLibelleBal);
    BalSit.SetAbrege ( FAbregeBal ) ;
    BalSit.SetExercice ( Exercice );
    BalSit.SetDateDebut( DateDeb );
    BalSit.SetDateFin ( DateFin );
    BalSit.SetDevise ( V_PGI.DevisePivot );
    // Ajout de la balance dans la liste
    FListeBalance.Add ( BalSit );
    { Création des lignes de détail de la balance }
    while not Q.Eof do
    begin
      if ((Q.FindField('G_COLLECTIF').AsString = 'X') and (FDebCreColl)) then
      begin
        Debit := Q.FindField('HB_DEBIT').AsFloat;
        Credit := Q.FindField('HB_CREDIT').AsFloat;
        DebitDev := Q.FindField('HB_DEBITDEV').AsFloat;
        CreditDev := Q.FindField('HB_CREDITDEV').AsFloat;
      end else
      begin
        if Q.FindField('HB_DEBIT').AsFloat >= Q.FindField('HB_CREDIT').AsFloat then
        begin
          Debit := Q.FindField('HB_DEBIT').AsFloat-Q.FindField('HB_CREDIT').AsFloat;
          Credit := 0;
          DebitDev := Q.FindField('HB_DEBITDEV').AsFloat-Q.FindField('HB_CREDITDEV').AsFloat;
          CreditDev := 0;
        end else
        begin
          Debit := 0;
          Credit := Q.FindField('HB_CREDIT').AsFloat-Q.FindField('HB_DEBIT').AsFloat;
          DebitDev := 0;
          CreditDev := Q.FindField('HB_CREDITDEV').AsFloat-Q.FindField('HB_DEBITDEV').AsFloat;
        end;
      end;
      BalSit.AjouteEcriture ( Q.FindField('HB_COMPTE1').AsString, '',
                          Debit, Credit,DebitDev, CreditDev);
      Q.Next;
    end;
  end;
  Ferme (Q);
end;

procedure TImpExpBds.ImportHistobal;
var Exercice : string;
    DateDeb, Datefin : TDateTime;
    stSelect, stWhere : string;
    Q : TQuery;
    bTous : boolean;
begin
  // Import depuis une balance de situation stockée dans Histobal
  stSelect := 'SELECT DISTINCT  HB_EXERCICE, HB_DATE1, HB_DATE2, HB_TYPEBAL FROM HISTOBAL';
  stWhere := ' WHERE HB_TYPEBAL="BDS"';
  if FFileName <> '' then
  begin
    Exercice := ReadTokenSt (FFileName);
    DateDeb := StrToDate(ReadTokenSt(FFilename));
    DateFin := StrToDate(ReadTokenSt(FFilename));
    stWhere := stWhere + ' AND HB_EXERCICE="'+Exercice+
      '" AND HB_DATE1="'+USDateTime(DateDeb)+'" AND HB_DATE2="'+USDateTime(DateFin)+'"';
    bTous := False;
  end else bTous := True;
  Q := OpenSQL (stSelect+stWhere,True);
  while not Q.Eof do
  begin
    if ImportBalanceHistobal ( Q.FindField('HB_EXERCICE').AsString,
          Q.FindField('HB_DATE1').AsDateTime, Q.FindField('HB_DATE2').AsDateTime, bTous) then
      Q.Next
    else break;
  end;
  Ferme (Q);
end;

procedure TImpExpBds.ImportPGI;
begin
  // Import au form at PGI
end;

procedure TImpExpBds.ImportSiscoII;
var F : TextFile;
  St : string;
  Gene, Auxi : string;
  bCollectif : boolean;
begin
  FLastError := 0;
  AssignFile ( F , FFileName);
  Reset (F);
  InitMove(100,'') ;
  while not Eof (F) do
  begin
    Readln (F, St);
    MoveCur(FALSE) ;
    if Copy (St,1,3) = '***' then continue
    else
    if Copy (St,1,2) = ' E' then // Enregistrement Entête
    begin
      // Création des balances
      if not CreationBalanceSiscoII ( St ) then break;
    end else
    if (St[11] = 'C') and (Copy (St,1,10)<>'}}}}}}}}}}') then  // Enregistrement Compte
    begin
      if not RecupereCompteSiscoII ( St, Gene, Auxi, True, bCollectif ) then break;
    end else
    if St[11] = 'b' then  // Enregistrement Détail
    begin
       if not AjouteLigneSiscoII ( St, Gene, Auxi, bCollectif ) then break;
    end;
  end;
  FiniMove ;
  CloseFile (F);
end;

procedure TImpExpBds.InsertDatabase;
var   i : integer;
      BalSit : TBalSit;
begin
  // Création des comptes
  FTGen.InsertDB(nil);
  FTAux.InsertDB(nil);
  // Création des balances
  for i := 0 to FListeBalance.Count - 1 do
  begin
    BalSit := FListeBalance.Items[i];
    BalSit.Enregistre;
  end;
  if (FListeBalance.Count = 0) then
  begin
    if (Assigned(FOnInformation)) then
    begin
      FLastError := ERR_NOBALANCE;
      FOnInformation(self, TraduireMemoire('Aucune balance restaurée.'), True);
    end;
  end;
end;

function TImpExpBds.LastErrorMsg: string;
begin
  Result := ERROR_MSG[fLastError];
end;

function TImpExpBds.RecupereCompteSiscoII(St: string; var Gene,
  Auxi: string; Creation: boolean; var bCollectif : boolean): boolean;
var bExiste : boolean;
    TG : TGeneral;
    TT : TTiers;
    Msg : string;
begin
  Auxi := ''; Gene := '';
  // si ligne sur compte auxiliaire
  if (St[1]='0') or (St[1]='F') or (St[1]='9') or (St[1]='C') or (St[64]='A') then
  begin
    Auxi := BourreLaDonc(Trim(Copy (St,1,Min(10,VH^.Cpta[fbAux].Lg))),fbAux);
    bExiste := Presence ('TIERS','T_AUXILIAIRE',Auxi) or (FTAux.FindFirst(['T_AUXILIAIRE'],[Auxi],False)<>nil);
    if not bExiste then
    begin
      if Creation then
      begin
        TT := TTiers.Create ( FTAux );
        TT.InitNouveau ( Gene, Auxi, True );
        TT.Libelle := Copy ( St, 39 , 25 );
        Gene := TT.GetValue('T_COLLECTIF');
        if Gene = '' then
        begin
          fLastError := ERR_PARAMSOC;
          Result := False;
          exit;
        end;
        if not Presence('GENERAUX','G_GENERAL',Gene) then
        begin
          fLastError := ERR_GENE_INEXISTANT;
          Result:= False;
          exit;
        end;
      end else
      begin
        fLastError := ERR_AUXI_INEXISTANT;
        Result:= False;
        exit;
      end;
    end else Gene := GetColonneSQL ('TIERS','T_COLLECTIF','T_AUXILIAIRE="'+Auxi+'"');
  end else
  begin
    Gene := BourreLaDonc(Trim(Copy (St,1,Min(10,VH^.Cpta[fbGene].Lg))),fbGene);
    bExiste := Presence ('GENERAUX','G_GENERAL',Gene) or (FTGen.FindFirst(['G_GENERAL'],[Gene], False)<>nil);
    if not bExiste then
    begin
      if Creation then
      begin
        TG := TGeneral.Create ( FTGen );
        TG.InitNouveau ( Gene );
        TG.Libelle := Copy ( St, 39 , 25 );
        TG.SetNature ( GeneNatureSisco2VersPGI (St[64]));
        bCollectif := TG.GetValue('G_COLLECTIF')='X';
      end else
      begin
        fLastError := ERR_GENE_INEXISTANT;
        Result:= False;
        exit;
      end;
    end else bCollectif := (GetColonneSQL('GENERAUX','G_COLLECTIF','G_GENERAL="'+Gene+'"')='X');
  end;
  if ((Gene <> '') and (Auxi <> '')) then Msg := 'Compte auxiliaire : '+Auxi
  else if Gene <> '' then Msg := TraduireMemoire('Compte général : '+Gene);
  if Assigned(FOnInformation) then FOnInformation(self, Msg , FALSE);
  Result := True;
end;


end.
