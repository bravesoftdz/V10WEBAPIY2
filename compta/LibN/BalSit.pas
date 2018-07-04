unit BalSit;

interface

uses Classes,
     Controls,
     SysUtils,
     uTOB,
     HEnt1,
     Hctrls,
     HMsgBox,

{$IFDEF EAGLSERVER}
{$ELSE}
     LookUp,
{$ENDIF}

{$IFDEF EAGLCLIENT}
{$ELSE}
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
     CritEdt;

type
  TBalSitInformation = procedure (Sender : TObject; Msg : string ) of object;

  RBalSitInfo = record
    Libelle : string;
    Abrege : string;
    Plan : string;
    Compte1Inf : string;
    Compte1Sup : string;
    Compte2Inf : string;
    Compte2Sup : string;
    DateInf : TDateTime;
    DateSup : TDateTime;
    CumExo : boolean;
    Exo : string;
    Ano : boolean;
    CumMois : boolean;
    CumPer : boolean;
    Etab : string;
    CumEtab : boolean;
    Devise : string;
    CumDev : boolean;
    TypeEcr : TEcrPCL;
    Axe : string;
    DebCreColl : boolean;
    TypeBalance : string;
  end;

  TBalSit = class
    private
      FCodeBal        : string;
      FLibelle        : string;
      FAbrege         : string;
      FTypeCumul      : string;
      FTypeBalance    : string;
      FExercice       : string;
      FDateDebut      : TDateTime;
      FDateFin        : TDateTime;
      FDevise         : string;
      FExiste         : boolean;
      FOnInformation  : TBalSitInformation;
      FModifie        : boolean;
      FNbEcriture     : integer;
      FDebCreColl: boolean;
      procedure InitDefaut;
      function FabriqRequeteBalAuto ( RBS : RBalSitInfo; var Compte1, Compte2, Prefixe, GroupBy : string ) : string;
      function FabriqRequeteBalEcr ( AvecLibelle : boolean ) : string;
      function  AffecteEnteteBalance ( RBS : RBalSitInfo ; Compte1, Compte2, Prefixe : string; var CodeBal : string ; Exercice, Etablissement, Devise,PeriodeMois : string) : TOB;
    public
      LEcr : TOB;
      constructor Create (CodeBal : string);
      destructor Destroy; override;
      procedure SetLibelle ( Libelle : string );
      procedure SetAbrege ( Abrege : string );
      procedure SetCumul ( TypeCumul : string );
      procedure SetExercice ( Exercice : string );
      procedure SetDevise ( Devise : string );
      procedure SetDateDebut ( DateDebut : TDateTime );
      procedure SetDateFin ( DateFin : TDateTime );
      procedure Enregistre;
      procedure Charge ( AvecLibelle : boolean );
      procedure Supprime;
      procedure GenereAuto(RBS : RBalSitInfo);
      procedure GenereManu(RBS : RBalSitInfo);
      procedure EnregistreBalanceAuto( RBS : RBalSitInfo; T : TOB; TobAno : Tob; Compte1, Compte2, Prefixe : string; Exercice, Etablissement, Devise, PeriodeMois : string);
      procedure AjouteEcriture (Compte1, Compte2 : string; Debit, Credit : double; DebitDev : double = 0; CreditDev : double = 0 );
      function  GetEcriture ( i : integer ) : TOB;
      procedure MajEcriture ( i : integer ; Debit, Credit : double );
      procedure SuppEcritureAZero;
      procedure TriParCompte;
      function  PresenceCompte ( Compte1, Compte2 : string ) : integer;
      function  SupprimeLigne ( Compte1, Compte2 : string ) : boolean;
    published
      property Existe     : boolean read FExiste;
      property CodeBal    : string read FCodeBal;
      property Libelle    : string read FLibelle write SetLibelle;
      property Abrege     : string read FAbrege write SetAbrege;
      property TypeCumul  : string read FTypeCumul write SetCumul;
      property TypeBalance: string read FTypeBalance write FTypeBalance;
      property Exercice   : string read FExercice write SetExercice;
      property DateDebut  : TDateTime read FDateDebut write SetDateDebut;
      property DateFin    : TDateTime read FDateFin write SetDateFin;
      property Devise     : string read FDevise write SetDevise;
      property NbEcriture : integer read FNbEcriture;
      property OnInformation : TBalSitInformation read FOnInformation write FOnInformation;
  end;


{$IFNDEF EAGLSERVER}
procedure LookUpBalSit ( E: TControl; TypeCumul, Exo, Etab : string);
{$ENDIF}

  // GCO - 06/09/2006 - Déplacement de la fonction pour compilation en Process Serveur
function CreationBDSDynamique( vStPeriodicite : string ) : Boolean;



implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  ULibExercice,
  {$ENDIF MODENT1}
  Ent1;


////////////////////////////////////////////////////////////////////////////////
// GCO - 06/09/2006 - Déplacement de la fonction pour compilation en Process Serveur
function CreationBDSDynamique( vStPeriodicite : string ) : Boolean;
var lPremMois, lPremAnnee, lNbBalance : word;
    lNbMois : word;
    i : integer;
    lDateFin : TDateTime;
    lJourFin  : Word;
    lMoisFIn  : Word;
    lAnneeFin : Word;
    lOffSet : integer;
    lStMoisFin : string;
    lStCodeBal : string;
    lStLib     : string;

    procedure _Creation;
    var lBS : TBalSit;
        lInfoBS : RBalSitInfo;
    begin
      lBS := TBalSit.Create( lStCodeBal );
      lInfoBS.Libelle        := 'B. de situation dynamique ' + lStLib + ' ' +
                                 FormatDatetime('mmmm', lDateFin) + ' ' +
                                 FormatDatetime('yyyy', lDateFin);
      lInfoBS.Abrege         := Copy( FormatDatetime('mmmm', lDateFin) + ' ' +
                                      FormatDatetime('yyyy', lDateFin) , 1, 17);
      lInfoBS.Plan           := 'GEN';
      lInfoBS.Compte1Inf     := '';
      lInfoBS.Compte1Sup     := '';
      lInfoBS.Compte2Inf     := '';
      lInfoBS.Compte2Sup     := '';
      lInfoBS.DateInf        := VH^.EnCours.Deb;
      lInfoBS.DateSup        := lDateFin;
      lInfoBS.CumExo         := False;
      lInfoBS.Exo            := VH^.EnCours.Code;
      lInfoBS.CumMois        := False;
      lInfoBS.CumPer         := False;
      lInfoBS.Etab           := '';
      lInfoBS.CumEtab        := False;
      lInfoBS.Devise         := V_Pgi.DevisePivot;
      lInfoBS.CumDev         := False;
      lInfoBs.Ano            := True;
      lInfoBS.TypeBalance    := 'DYN';
      lInfoBS.TypeEcr[Reel]  := True;
      lInfoBS.TypeEcr[Simu]  := False;
      lInfoBS.TypeEcr[Situ]  := False;
      lInfoBS.TypeEcr[Previ] := False;
      lInfoBS.TypeEcr[Revi]  := False;
      lInfoBS.TypeEcr[Ifrs]  := False;
      lInfoBS.Axe            := '';
      lInfoBS.DebCreColl     := False;

      lBS.GenereAuto ( lInfoBS );
      lBS.Free;
    end;

begin
  Result := False;

  lNBBalance := 0;
  lOffSet := 0;

  NombreMois( VH^.Encours.Deb , VH^.Encours.Fin , lPremMois, lPremAnnee, lNbMois);
  if vStPeriodicite = 'M' then
  begin
    lNbBalance := lNbMois;
    lOffSet := 1;
    lStLib := 'mensuelle';
  end
  else
  if vStPeriodicite = 'T' then
  begin
    lNbBalance := lNbMois div 3;
    lOffSet := 3;
    lStLib := 'trimestrielle';
  end
  else
  if vStPeriodicite = 'S' then
  begin
    lNbBalance := lNbMois div 6;
    lOffSet := 6;
    lStLib := 'semestrielle';
  end
  else
  if vStPeriodicite = 'A' then
  begin
    lNbBalance := 1;
    lStLib := 'annuelle';
  end;

  // GCO - 12/09/2006 - FQ 18782
  for i := 1 to lNbBalance do
  begin
    if (lNbBalance = 1) and (vStPeriodicite = 'A') then
      lDateFin := FinDeMois(VH^.Encours.Fin)
    else
      lDateFin := FinDeMois(PlusMois(VH^.Encours.Deb, i * lOffset - 1));

    DecodeDate( lDateFin, lAnneeFin, lMoisFin, lJourFin );

    lStMoisFin := IntToStr(lMoisFin);

    if Length(lStMoisFin) = 1 then
      lStMoisFin := '0' + lStMoisFin;

    lStCodeBal := 'BDS@' + IntToStr(lAnneeFin) + lStMoisFin + vStPeriodicite;

    if not ExisteSQL('SELECT BSI_CODEBAL FROM CBALSIT WHERE BSI_CODEBAL = "' + lStCodeBal + '"') then
      _Creation;
  end;
end;

////////////////////////////////////////////////////////////////////////////////

{ TBalSit }

procedure TBalSit.Charge ( AvecLibelle : boolean );
var Q : TQuery;
    stRequete : string;
begin
  if LEcr = nil then exit;
  LEcr := TOB.Create ('CBALSIT', nil , -1);
  LEcr.SelectDB ('"'+FCodeBal+'"',nil);
  FTypeCumul := LEcr.GetValue('BSI_TYPECUM');
  FTypeBalance := lEcr.GetString('BSI_TYPEBAL');
  FLibelle := LEcr.GetValue('BSI_LIBELLE');
  FAbrege := LEcr.GetValue('BSI_ABREGE');
  FDebCreColl := LEcr.GetValue('BSI_DEBCRECOLL')='X';
  FDateDebut := LEcr.GetValue('BSI_DATE1');
  FDateFin := LEcr.GetValue('BSI_DATE2');
  stRequete := FabriqRequeteBalEcr (AvecLibelle);
  Q := OpenSQL (stRequete,True);
  try
    LEcr.LoadDetailDB('CBALSITECR','','',Q,True,True);
  finally
    Ferme (Q);
  end;
  FNbEcriture := LEcr.Detail.Count;
end;

constructor TBalSit.Create(CodeBal: string);
begin
  FNbEcriture := 0;
  FCodeBal := CodeBal;
  LEcr := TOB.Create ('CBALSIT',nil, -1);
  FExiste := ExisteSQL('SELECT BSI_CODEBAL FROM CBALSIT WHERE BSI_CODEBAL="'+FCodeBal+'"');
  FModifie := FExiste;
  if not FExiste then InitDefaut;
end;

destructor TBalSit.Destroy;
begin
  LEcr.Free;
  inherited;
end;

procedure TBalSit.Enregistre;
begin
  ExecuteSQL ('DELETE FROM CBALSITECR WHERE BSE_CODEBAL="'+FCodeBal+'"');
  ExecuteSQL ('DELETE FROM CBALSIT WHERE BSI_CODEBAL="'+FCodeBal+'"');
  LEcr.SetAllModifie(True);
  // On n'enregistre pas la balance si aucune ligne
  // FQ 20559 - CA - 18/06/2007 - Si balance dynamique vide, on enregistre
  if ((LEcr.Detail.Count > 0) or (TypeBalance='DYN')) then LEcr.InsertDB ( nil, True );
end;

procedure TBalSit.GenereAuto(RBS : RBalSitInfo);
var   QBal   :   TQuery;
      stReqPrinc, stRequete  :   string;
      T, TExercice, TEtablissement, TDevise : TOB;
      TAno : Tob;
      LMois : TStringList;
      i, ex, et, de, mo : integer;
      Compte1, Compte2, Prefixe, Cle : string;
      St, stReqExo, stReqEtab, stReqMois, stGroupBy : string;
      Exercice, Etablissement, Devise, Periode : string;
      DateExo : TExoDate;
      MM, AA, NbMois : Word;
      DD : TDateTime;
begin
  TAno := nil;

  // Génération de la requête principale
  stReqPrinc := FabriqRequeteBalAuto(RBS, Compte1, Compte2, Prefixe, stGroupBy);

  // Chargement des critères
  if (RBS.CumExo) or (RBS.Exo = '') then Cle := '' else Cle := '"'+RBS.Exo+'"';
  TExercice := TOB.Create ('', nil, -1);
  TExercice.LoadDetailDB('EXERCICE',Cle,'',nil,False);
  if (RBS.CumEtab) or (RBS.Etab='') then Cle := '' else Cle := '"'+RBS.Etab+'"';
  TEtablissement := TOB.Create ('', nil, -1);
  TEtablissement.LoadDetailDB('ETABLISS',Cle,'',nil,False);
  if (RBS.CumDev) or (RBS.Devise='') then Cle := '' else Cle := '"'+RBS.Devise+'"';
  TDevise := TOB.Create ('', nil, -1);
  TDevise.LoadDetailDB('DEVISE',Cle,'',nil,False);
  // Chargement de la liste des période mois
  if RBS.CumMois then
  begin
    LMois := TStringList.Create;
    for ex :=0 to TExercice.Detail.Count - 1 do
    begin
      DateExo.Deb := TExercice.Detail[ex].GetValue('EX_DATEDEBUT');
      DateExo.Fin := TExercice.Detail[ex].GetValue('EX_DATEFIN');
      NbMois:=0 ;
      NOMBREPEREXO(DateExo,MM,AA,NbMois) ;
      for i:=0 to NbMois-1 do
      begin
        DD := PlusMois(DateExo.Deb,i) ;
        LMois.Add (FormatDateTime('yyyymm',DD));
      end;
    end;
  end else LMois := nil;

  for ex:=0 to TExercice.Detail.Count - 1 do
  begin
    Exercice := TExercice.Detail[ex].GetValue('EX_EXERCICE');
    St := ' AND '+Prefixe+'_EXERCICE="'+Exercice+'"';
    if RBS.CumExo then stReqExo := stReqPrinc + St
    else stReqExo := stReqPrinc;
    for et:=0 to TEtablissement.Detail.Count - 1 do
    begin
      Etablissement := TEtablissement.Detail[et].GetValue('ET_ETABLISSEMENT');
      St := ' AND '+Prefixe+'_ETABLISSEMENT="'+Etablissement+'"';
      if RBS.CumEtab then stReqEtab := stReqExo + St
      else stReqEtab := stReqExo;
      for de:=0 to TDevise.Detail.Count - 1 do
      begin
        Devise := TDevise.Detail[de].GetValue('D_DEVISE');
        St := ' AND '+Prefixe+'_DEVISE="'+Devise+'"';
        if RBS.CumDev then stRequete := stReqEtab + St
        else StRequete := stReqEtab;
        // Génération de la balance avec cumul par mois
        if (RBS.CumMois) and (LMois <> nil) then
        begin
          for mo:=0 to LMois.Count - 1 do
          begin
            St := ' AND '+Prefixe+'_PERIODE='+LMois[mo];
            stReqMois := stRequete+St;
            QBal := OpenSQL( stReqMois + stGroupBy, True ) ;
            if not QBal.Eof then
            begin
              T := TOB.Create ('', nil, -1);
              T.LoadDetailDB ('','','',QBal,False,True);

              TAno := Tob.Create('', nil,-1);
              TAno.LoadDetailFromSQL( StRequete + ' AND ((E_ECRANOUVEAU = "H") OR (E_ECRANOUVEAU = "OAN")) ' +
                                      stGroupBy, False);

              EnregistreBalanceAuto( RBS, T, TAno, Compte1, Compte2, Prefixe, Exercice, Etablissement, Devise,LMois[mo]);
            end;
            Ferme (QBal);
          end;
        end else
        // Génération de la balance sans cumul par mois
        begin
          QBal := OpenSQL( stRequete + stGroupBy, True ) ;
          if ((not QBal.Eof) or (RBS.TypeBalance='DYN')) then
          begin
            T := TOB.Create ('', nil, -1);
            T.LoadDetailDB ('','','',QBal,False,True);

            TAno := Tob.Create('', nil,-1);
            TAno.LoadDetailFromSQL( StRequete + ' AND ((E_ECRANOUVEAU = "H") OR (E_ECRANOUVEAU = "OAN")) ' +
                                    stGroupBy, False);
            EnregistreBalanceAuto( RBS, T, TAno, Compte1, Compte2, Prefixe, Exercice, Etablissement, Devise, '');
          end;
          Ferme (QBal);
        end;
        if not RBS.CumDev then break;
      end;
      if not RBS.CumEtab then break;
    end;
    if not RBS.CumExo then break;
  end;

  if RBS.CumMois then LMois.Free;
  TExercice.Free;
  TEtablissement.Free;
  TDevise.Free;

  if assigned(TAno) then TAno.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 16/11/2001
Modifié le ... :   /  /
Description .. : Fabrication de la requête principale pour génération d'une
Suite ........ : balance automatique
Mots clefs ... : SITUATION;BALANCE
*****************************************************************}
function TBalSit.FabriqRequeteBalAuto(RBS: RBalSitInfo; var Compte1, Compte2, Prefixe, GroupBy : string): string;
var SQL, Etabl, Dev, Table, DateSel : String ;
    AuxAna, Montant, Where   : String ;
    Exo2, TypEcr, Periode,ChampColl, Jointure : String ;
    stSelect, CompteAux : string;
begin
  SQL := ''; Jointure := '';ChampColl := ''; CompteAux := '';

  if RBS.Plan = 'GEN' then
  begin
    Compte1:='E_GENERAL';Compte2:='';Prefixe:='E' ; ChampColl:= ',G_COLLECTIF'; Jointure := ' LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL';
    if RBS.DebCreColl then CompteAux :=',E_AUXILIAIRE';
  end else
  if RBS.Plan = 'TIE' then
  begin Compte1:='E_AUXILIAIRE'; Compte2:=''; Prefixe:='E'; end else
  if RBS.Plan = 'ANA' then
  begin Compte1:='Y_SECTION'; Compte2:=''; Prefixe:='Y'; end else
  if RBS.Plan = 'G/A' then
  begin Compte1:='Y_GENERAL'; Compte2:='Y_SECTION'; Prefixe:='Y' ; Where:=' AND '+Compte2+'<>""' ; end else
  if RBS.Plan = 'G/T' then
  begin Compte1:='E_GENERAL'; Compte2:='E_AUXILIAIRE'; Prefixe:='E'; Where:=' AND '+Compte2+'<>""' ; end;

  if AnsiPos('*',RBS.Compte1Inf)>0 Then
    Where:=Where+' AND '+Compte1+' LIKE "'+Copy(RBS.Compte1Inf,1,AnsiPos('*',RBS.Compte1Inf)-1)+'%"'
  else if (RBS.Compte1Inf<>'') Then
  begin
    if RBS.Compte1Sup <>'' Then Where:=Where+' AND '+Compte1+'>="'+RBS.Compte1Inf+'"'+' AND '+Compte1+'<="'+RBS.Compte1Sup+'"' ;
    if RBS.Compte1Sup =''  Then Where:=Where+' AND '+Compte1+'>="'+RBS.Compte1Inf+'"' ;
  end;
  If AnsiPos('*',RBS.Compte2Inf)=0 Then
  begin
    if RBS.Compte2Inf<>'' Then Where:=Where+' AND '+Compte2+'>="'+RBS.Compte2Inf+'"' ;
    if RBS.Compte2Sup<>'' Then Where:=Where+' AND '+Compte2+'<="'+RBS.Compte2Sup+'"' ;
  end else Where:=Where+' AND '+Compte2+' LIKE "'+Copy(RBS.Compte1Inf,1,AnsiPos('*',RBS.Compte1Inf)-1)+'%"' ;

  if RBS.Exo='' then if RBS.CumExo then Exo2:=', '+Prefixe+'_EXERCICE' else Exo2:='' ;
  if RBS.Exo<>'' then Where:=Where+' AND '+Prefixe+'_EXERCICE'+'="'+RBS.Exo+'"' ;

  if RBS.CumMois then Periode:=', '+Prefixe+'_PERIODE'
  else Periode:='' ;

  if RBS.Etab='' then if RBS.CumEtab then Etabl:=', '+Prefixe+'_ETABLISSEMENT' else Etabl:='' ;
  if RBS.Etab<>''then Where:=Where+' AND '+Prefixe+'_ETABLISSEMENT'+'="'+RBS.Etab+'"' ;

  if RBS.Devise='' then if RBS.CumDev then Dev:=', '+Prefixe+'_DEVISE' Else Dev:='';
  if RBS.Devise<>'' then Where:=Where+' AND '+Prefixe+'_DEVISE'+'="'+RBS.Devise+'"' ;

  if RBS.TypeEcr[Reel]=True then TypEcr:=TypEcr+' OR '+Prefixe+'_QUALIFPIECE="N"' ;
  if RBS.TypeEcr[Simu]=True then TypEcr:=TypEcr+' OR '+Prefixe+'_QUALIFPIECE="S"' ;
  if RBS.TypeEcr[Situ]=True then TypEcr:=TypEcr+' OR '+Prefixe+'_QUALIFPIECE="U"' ;
  if RBS.TypeEcr[Previ]=True then TypEcr:=TypEcr+' OR '+Prefixe+'_QUALIFPIECE="P"' ;
  if RBS.TypeEcr[Revi]=True then TypEcr:=TypEcr+' OR '+Prefixe+'_QUALIFPIECE="R"' ;   // FQ 17107 : SBO 03/01/2006 : Ajout des écritures de révision...
  if RBS.TypeEcr[Ifrs]=True then TypEcr:=TypEcr+' OR '+Prefixe+'_QUALIFPIECE="I"' ;
  if not RBS.Ano then TypEcr:=TypEcr+' AND ('+Prefixe+'_ECRANOUVEAU<>"H" AND '+Prefixe+'_ECRANOUVEAU<>"OAN") ' ;

  Delete(TypEcr,1,4) ;
  Where:=Where+' AND ('+TypEcr+')' ;

  if Prefixe='Y' Then Where:=Where+' AND Y_AXE="'+RBS.Axe+'"' ;
  If Prefixe='E' Then Table:='ECRITURE' Else Table:='ANALYTIQ' ;

  If Compte2<>'' Then AuxAna:=', '+Compte2 ;

  Montant:=', SUM('+Prefixe+'_DEBIT) DEBIT, SUM('+Prefixe+'_CREDIT) CREDIT, ' ;
  Montant:=Montant+'SUM('+Prefixe+'_DEBITDEV) DEBITDEV, SUM('+Prefixe+'_CREDITDEV) CREDITDEV' ;
  DateSel:=Prefixe+'_DATECOMPTABLE' ;

  SQL:='SELECT '+Compte1+CompteAux+AuxAna+Exo2+Periode+stSelect+Montant+ChampColl+' FROM '+Table+Jointure+' WHERE '+DateSel+'>="' ;
  SQL:=SQL+USDateTime(RBS.DateInf)+'" AND '+DateSel+'<="'+USDateTime(RBS.DateSup)+'"' ;
  if RBS.Plan = 'TIE' then SQL := SQL + ' AND '+Prefixe+'_AUXILIAIRE<>"" ';
  SQL := SQL+Where;
  GroupBy := ' GROUP BY '+Compte1+CompteAux+AuxAna+Exo2+Periode+Etabl+Dev+ChampColl+' ORDER BY '+Compte1+AuxAna+Exo2+Periode+Etabl+Dev ;
  Result := SQL;
end;

procedure TBalSit.EnregistreBalanceAuto( RBS : RBalSitInfo; T : TOB; TobAno : Tob; Compte1, Compte2, Prefixe : string ; Exercice, Etablissement, Devise, PeriodeMois : string);
var   TBal, TEcr : TOB;
      CodeBal : string;
      i , k , l: integer;
      ValCompte1, ValCompte2, LibCompte1, LibCompte2, Msg : string;
      Debit, Credit, DebitDev, CreditDev : double;
      DebitAno, CreditAno, DebitAnoDev, CreditAnoDev : double;
      bCollectif : boolean;
begin
  TBal := AffecteEnteteBalance ( RBS , Compte1, Compte2, Prefixe , CodeBal, Exercice, Etablissement, Devise,PeriodeMois);
  k := -1;
  for i:=0 to T.Detail.Count-1 do
  begin
    if i <= k then continue else k:=-1;  // pour se repositionner sur la bonne écriture en DebCreColl
    ValCompte1 := T.Detail[i].GetValue(Compte1);
    if Compte2 <> '' then ValCompte2 := T.Detail[i].GetValue(Compte2);
    Msg := LibCompte1+' : '+ValCompte1;
    if LibCompte2 <> '' then Msg := Msg+' - '+LibCompte2+' : '+ValCompte2;
    if assigned (FOnInformation) then FOnInformation(Self,Msg);

    TEcr := TOB.Create ('CBALSITECR', TBal, -1);
    TEcr.PutValue('BSE_CODEBAL', CodeBal);
    TEcr.PutValue('BSE_COMPTE1', ValCompte1);

    if Compte2<>'' then TEcr.PutValue('BSE_COMPTE2', ValCompte2);

    if RBS.Plan = 'GEN' then
      bCollectif := T.Detail[i].GetValue('G_COLLECTIF')='X'
    else
      bCollectif := False;

    DebitAno     := 0;
    CreditAno    := 0;
    DebitAnoDev  := 0;
    CreditAnoDev := 0;

    if (bCollectif) and (RBS.DebCreColl) then
    begin
      // Cumuls par auxiliaire pour soldes des débiteurs et créditeurs
      k := i;
      Debit := 0; Credit := 0; DebitDev := 0; CreditDev := 0;
      while T.Detail[k].GetValue(Compte1)=ValCompte1 do
      begin
        if T.Detail[k].GetValue('DEBIT') > T.Detail[k].GetValue('CREDIT') then
        begin
          Debit := Debit + T.Detail[k].GetValue('DEBIT') - T.Detail[k].GetValue('CREDIT');
          DebitDev := DebitDev + T.Detail[k].GetValue('DEBITDEV') - T.Detail[k].GetValue('CREDITDEV');
        end
        else
        begin
          Credit := Credit + T.Detail[k].GetValue('CREDIT') - T.Detail[k].GetValue('DEBIT');
          CreditDev := CreditDev + T.Detail[k].GetValue('CREDITDEV') - T.Detail[k].GetValue('DEBITDEV');
        end;

        if ((k = T.Detail.Count-1) or (T.Detail[k+1].GetValue(Compte1)<>ValCompte1)) then
          break
        else
          Inc (k,1);
      end;

      TEcr.PutValue('BSE_DEBIT', Debit);
      TEcr.PutValue('BSE_CREDIT', Credit);


      // Même calcul dans la TobANo
      l := 0;
      if TobAno.Detail.Count > 0 then // GCO - 12/05/2006 - FQ 17894
      begin
        while TobAno.Detail[l].GetString(Compte1) = valCompte1 do
        begin
          if TobAno.Detail[l].GetDouble('DEBIT') > TobAno.Detail[l].GetDouble('CREDIT') then
          begin
            DebitAno := DebitAno + TobAno.Detail[l].GetDouble('DEBIT') - TobAno.Detail[l].GetDouble('CREDIT');
            DebitAnoDev := DebitAnoDev + TobAno.Detail[l].GetDouble('DEBITDEV') - TobAno.Detail[l].GetDouble('CREDITDEV');
          end
          else
          begin
            CreditAno := CreditAno + TobAno.Detail[l].GetDouble('CREDIT') - TobAno.Detail[l].GetDouble('DEBIT');
            CreditAnoDev := CreditAnoDev + TobAno.Detail[l].GetDouble('CREDITDEV') - TobAno.Detail[l].GetDouble('DEBITDEV');
          end;

          if ((l = TobAno.Detail.Count-1) or (TobAno.Detail[l+1].GetString(Compte1) <> valCompte1)) then
            break
          else
            Inc(l,1);
        end;
      end;

      TEcr.PutValue('BSE_DEBITANO', DebitAno);
      TEcr.PutValue('BSE_CREDITANO', CreditAno); 

    end
    else
    begin
      Debit := T.Detail[i].GetValue('DEBIT');
      Credit := T.Detail[i].GetValue('CREDIT');
      DebitDev := T.Detail[i].GetValue('DEBITDEV');
      CreditDev := T.Detail[i].GetValue('CREDITDEV');

      if (TobAno <> nil) and (TobAno.Detail.Count > 0) and
         (TobAno.Detail[0].GetString(Compte1) = valCompte1) and
         (TobAno.Detail[0].GetString(Compte2) = valCompte2) then
      begin
        DebitAno     := TobAno.Detail[0].GetDouble('DEBIT');
        CreditAno    := TobAno.Detail[0].GetDouble('CREDIT');
        DebitAnoDev  := TobAno.Detail[0].GetDouble('DEBITDEV');
        CreditAnoDev := TobAno.Detail[0].GetDouble('CREDITDEV');
        TobAno.Detail[0].Free;
      end;

      if (Debit > Credit ) then
        TEcr.PutValue('BSE_DEBIT', Debit-Credit)
      else
        TEcr.PutValue('BSE_CREDIT', Credit-Debit);

      if (DebitAno > CreditAno) then
        TEcr.PutValue('BSE_DEBITANO', DebitAno - CreditAno)
      else
        TEcr.PutValue('BSE_CREDITANO', CreditAno - DebitAno);

    end;

    if ( RBS.Devise='' ) and ( not RBS.CumDev ) Then
    begin
      TEcr.PutValue('BSE_DEBITDEV',TEcr.GetValue('BSE_DEBIT')) ;
      TEcr.PutValue('BSE_CREDITDEV',TEcr.GetValue('BSE_CREDIT')) ;

      TEcr.PutValue('BSE_DEBITANODEV', TEcr.GetValue('BSE_DEBITANO'));
      TEcr.PutValue('BSE_CREDITANODEV', TEcr.GetValue('BSE_CREDITANO'));
    end
    else
    begin
      if (bCollectif) and (RBS.DebCreColl) then
      begin
        TEcr.PutValue('BSE_DEBITDEV', DebitDev);
        TEcr.PutValue('BSE_CREDITDEV', CreditDev);

        TEcr.PutValue('BSE_DEBITANODEV', DebitAnoDev);
        TEcr.PutValue('BSE_CREDITANODEV', CreditAnoDev);
      end
      else
      begin
        if (DebitDev > CreditDev ) then
          TEcr.PutValue('BSE_DEBITDEV', DebitDev - CreditDev)
        else
          TEcr.PutValue('BSE_CREDITDEV', CreditDev - DebitDev);

        if (DebitAnoDev > CreditAnoDev) then
          TEcr.PutValue('BSE_DEBITANODEV', DebitAnoDev - CreditAnoDev)
        else
          TEcr.PutValue('BSE_CREDITANODEV', CreditAnoDev - DebitAnoDev);
      end;
    end;
  end;

  if assigned (FOnInformation) then FOnInformation(Self,TraduireMemoire('Mise à jour de la base en cours ...'));
  TBal.InsertDB(nil);
  TBal.Free;
  if assigned (FOnInformation) then FOnInformation(Self,TraduireMemoire('Enregistrement terminé.'));
end;

procedure TBalSit.Supprime;
begin
  ExecuteSQL ('DELETE FROM CBALSIT WHERE BSI_CODEBAL="'+FCodeBal+'"');
  ExecuteSQL ('DELETE FROM CBALSITECR WHERE BSE_CODEBAL="'+FCodeBal+'"');
end;

procedure TBalSit.AjouteEcriture(Compte1, Compte2: string; Debit,
  Credit: double; DebitDev : double = 0; CreditDev : double = 0);
var T : TOB;
begin
  T := TOB.Create ('CBALSITECR', LEcr,-1);
  if (T <> nil) then
  begin
    T.PutValue ( 'BSE_CODEBAL', FCodeBal );
    T.PutValue ( 'BSE_COMPTE1', Compte1 );
    T.PutValue ( 'BSE_COMPTE2', Compte2 );
    T.PutValue ( 'BSE_DEBIT', Debit );
    T.PutValue ( 'BSE_CREDIT', Credit );
    if DebitDev = 0 then T.PutValue ( 'BSE_DEBITDEV', Debit )
    else T.PutValue ( 'BSE_DEBITDEV', DebitDev );
    if CreditDev = 0 then T.PutValue ( 'BSE_CREDITDEV', Credit )
    else T.PutValue ( 'BSE_CREDITDEV', CreditDev );
    FModifie := True;
    Inc (FNbEcriture , 1);
  end;
end;

procedure TBalSit.SetCumul(TypeCumul: string);
begin
  FModifie := TypeCumul <> FTypeCumul;
  FTypeCumul := TypeCumul;
  LEcr.PutValue('BSI_TYPECUM', FTypeCumul);
end;

procedure TBalSit.SetDateDebut(DateDebut: TDateTime);
begin
  FModifie := DateDebut <> FDateDebut;
  FDateDebut := DateDebut;
  LEcr.PutValue('BSI_DATE1', FDateDebut);
end;

procedure TBalSit.SetDateFin(DateFin: TDateTime);
begin
  FModifie := DateFin <> FDateFin;
  FDateFin := DateFin;
  LEcr.PutValue('BSI_DATE2', FDateFin);
end;

procedure TBalSit.SetDevise(Devise: string);
begin
  FModifie := Devise <> FDevise;
  FDevise := Devise;
  LEcr.PutValue('BSI_DEVISE', FDevise);
end;

procedure TBalSit.SetExercice(Exercice: string);
begin
  FModifie := Exercice <> FExercice;
  FExercice := Exercice;
  LEcr.PutValue('BSI_EXERCICE', FExercice);
end;

procedure TBalSit.SetLibelle(Libelle: string);
begin
  FModifie := Libelle <> FLibelle;
  FLibelle := Libelle;
  LEcr.PutValue('BSI_LIBELLE', FLibelle);
end;

procedure TBalSit.SetAbrege(Abrege: string);
begin
  FModifie := Abrege <> FAbrege;
  FAbrege := Abrege;
  LEcr.PutValue('BSI_ABREGE', FAbrege);
end;

procedure TBalSit.InitDefaut;
begin
  LEcr.PutValue ('BSI_CODEBAL',FCodeBal );
  LEcr.PutValue ('BSI_TYPEBAL', 'BDS' );
  LEcr.PutValue ('BSI_PLAN', 'O' );
  LEcr.PutValue ('BSI_AXE', 'A1' );
  LEcr.PutValue ('BSI_ETABLISSEMENT', '...' );
end;

function TBalSit.FabriqRequeteBalEcr ( AvecLibelle : boolean ) : string;
var St, stLibelle, stJointure, stOrder : string;
begin
  stLibelle := ''; stJointure := '';
  if AvecLibelle then
  begin
    if FTypeCumul='GEN' then
    begin
      stLibelle := ',G_LIBELLE LIBELLE';
      stJointure := ' LEFT JOIN GENERAUX ON (G_GENERAL=BSE_COMPTE1) ';
      stOrder := ' ORDER BY BSE_COMPTE1';
    end else if FTypeCumul='G/T' then
    begin
      stLibelle := ',T_LIBELLE LIBELLE ';
      stJointure := ' LEFT JOIN TIERS ON (T_AUXILIAIRE=BSE_COMPTE2) ';
      stOrder := ' ORDER BY BSE_COMPTE1, BSE_COMPTE2';
    end else if FTypeCumul='G/A' then
    begin
      stLibelle := ',S_LIBELLE LIBELLE ';
      stJointure := ' LEFT JOIN SECTION ON (S_SECTION=BSE_COMPTE2) ';
      stOrder := ' ORDER BY BSE_COMPTE1, BSE_COMPTE2';
    end else if FTypeCumul='ANA' then
    begin
      stLibelle := ',S_LIBELLE LIBELLE ';
      stJointure := ' LEFT JOIN SECTION ON (S_SECTION=BSE_COMPTE1) ';
      stOrder := ' ORDER BY BSE_COMPTE1';
    end else if FTypeCumul='TIE' then
    begin
      stLibelle := ',T_LIBELLE LIBELLE ';
      stJointure := ' LEFT JOIN TIERS ON (T_AUXILIAIRE=BSE_COMPTE1) ';
      stOrder := ' ORDER BY BSE_COMPTE1';
    end;
  end;
  St := 'SELECT CBALSITECR.* '+stLibelle+' FROM CBALSITECR '+stJointure+
        ' WHERE BSE_CODEBAL="'+FCodeBal+'"'+ stOrder;
  Result := St;
end;

function TBalSit.GetEcriture(i: integer): TOB;
begin
  if i < LEcr.Detail.Count then
    Result := LEcr.Detail[i]
  else Result := nil;
end;

procedure TBalSit.MajEcriture(i: integer; Debit, Credit: double);
var T : TOB;
begin
  if i < LEcr.Detail.Count then
  begin
    T := LEcr.Detail[i];
    if (T.GetValue ('BSE_DEBIT') <> Debit) or (T.GetValue ('BSE_CREDIT') <> Credit) then
    begin
      T.PutValue('BSE_DEBIT',Debit);
      T.PutValue('BSE_CREDIT',Credit);
      FModifie := True;
    end;
  end;
end;

procedure TBalSit.GenereManu(RBS: RBalSitInfo);
var TBal, T : TOB;
  Compte1, Compte2, Requete, CodeBal : string;
  Q : TQuery;
begin
  if RBS.Plan = 'GEN' then
  begin
    Compte1 := 'G_GENERAL';
    Compte2 := '';
    Requete := 'SELECT G_GENERAL FROM GENERAUX';
  end else
  if RBS.Plan = 'TIE' then
  begin
    Compte1 := 'T_AUXILIAIRE';
    Compte2 := '';
    Requete := 'SELECT T_AUXILIAIRE FROM TIERS';
  end else
  if RBS.Plan = 'ANA' then
  begin
    Compte1 := 'S_SECTION';
    Compte2 := '';
    Requete := 'SELECT S_SECTION FROM SECTION';
  end else
  if RBS.Plan = 'G/T' then
  begin
    Compte1 := 'G_GENERAL';
    Compte2 := 'T_AUXILIAIRE';
    Requete := 'SELECT G_GENERAL,T_AUXILIAIRE FROM TIERS LEFT JOIN GENERAUX ON T_COLLECTIF=G_GENERAL';
  end else
  if RBS.Plan = 'G/A' then
  begin
//    Compte1 := 'G_GENERAL';
//    Compte2 := 'S_SECTION';
  MessageAlerte ('Non supporté.');
  exit;
  end;
  TBal := AffecteEnteteBalance (RBS,Compte1, Compte2, '', CodeBal,RBS.Exo,RBS.Etab,RBS.Devise, '');
  Q := OpenSQL (Requete,True);
  while not Q.Eof do
  begin
    T := TOB.Create ('CBALSITECR',TBal,-1);
    T.PutValue ('BSE_CODEBAL',CodeBal);
    T.PutValue ('BSE_COMPTE1',Q.FindField(Compte1).AsString);
    if Compte2 <> '' then
      T.PutValue ('BSE_COMPTE2',Q.FindField(Compte2).AsString);
    Q.Next;
  end;
  Ferme (Q);
  if TBal.Detail.Count > 0 then
    TBal.InsertDB(nil,True);
  TBal.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 16/11/2001
Modifié le ... :   /  /
Description .. : Crée l'entête de la balance de situation (CBALSIT)
Mots clefs ... : SITUATION;CREATION;ENTETE
*****************************************************************}
function TBalSit.AffecteEnteteBalance(RBS : RBalSitInfo; Compte1, Compte2, Prefixe : string; var CodeBal : string;  Exercice, Etablissement, Devise, PeriodeMois : string): TOB;
var TBal : TOB;
    LibCompte1, LibCompte2 : string;
    LibBal : string;
    QExo : TQuery;
    a,m,j : Word;
    DatePer : TDateTime;
    cpt : integer;
begin
  TBal := TOB.Create ('CBLASIT', nil, - 1);
  if TBal <> nil then
  begin
    // Affectation des libellés
    LibCompte2 := '';
    if RBS.Plan = 'GEN' then LibCompte1 := TraduireMemoire ('Général')
    else if RBS.Plan = 'TIE' then LibCompte1 := TraduireMemoire ('Auxiliaire')
    else if RBS.Plan = 'ANA' then LibCompte1 := TraduireMemoire ('Section')
    else
    begin
      LibCompte1 := TraduireMemoire ('Général');
      if RBS.Plan = 'G/T' then LibCompte2 := TraduireMemoire ('Auxiliaire');
      if RBS.Plan = 'G/A' then LibCompte2 := TraduireMemoire ('Section');
    end;

    // Calcul du code de la balance
    CodeBal := FCodeBal;
    if RBS.CumExo then  CodeBal := CodeBal + Exercice;
    if RBS.CumEtab then  CodeBal := CodeBal + Etablissement;
    if RBS.CumDev then  CodeBal := CodeBal + Devise;
    if RBS.CumMois then CodeBal := CodeBal + PeriodeMois;
    cpt := 0;

    // Si le code existe déjà, on ajoute un indice à la suite
    while ExisteSQL('SELECT * FROM CBALSIT WHERE BSI_CODEBAL="'+CodeBal+'"') do
    begin
      CodeBal := CodeBal + '_'+IntToStr(cpt);
      Inc ( cpt , 1 );
    end;

    // Calcul du libellé de la balance
    LibBal := RBS.Libelle;
    if RBS.CumExo then  LibBal := LibBal + ' - Exo : '+Exercice;
    if RBS.CumEtab then  LibBal := LibBal + ' - Etab : '+Etablissement;
    if RBS.CumDev then  LibBal := LibBal + ' - Dev : '+Devise;
    if RBS.CumMois then  LibBal := LibBal + ' - Mois : '+Copy(PeriodeMois,5,2)+'/'+Copy(PeriodeMois,1,4);

    // 1 - Création de la ligne d'entête de la balance (Table CBALSIT)
    TBal := TOB.Create ('CBALSIT',nil,-1);
    TBal.PutValue('BSI_CODEBAL', CodeBal);
    TBal.PutValue('BSI_LIBELLE', LibBal);
    TBal.PutValue('BSI_ABREGE', RBS.Abrege);
    if RBS.DebCreColl then TBal.PutValue('BSI_DEBCRECOLL', 'X');
    TBal.PutValue('BSI_TYPEBAL', 'BDS');
    TBal.PutValue('BSI_TYPECUM', RBS.Plan);
    // Si cumul par exercice (sans cumul par période)
    if (RBS.Exo = '') and ( RBS.CumExo ) and ( not (RBS.CumMois) ) then
    begin
      QExo := OpenSQL('SELECT EX_DATEDEBUT, EX_DATEFIN FROM EXERCICE WHERE EX_EXERCICE="'+Exercice+'"',TRUE) ;
      if not QExo.Eof then
      begin
        TBal.PutValue('BSI_DATE1',QExo.FindField('EX_DATEDEBUT').AsDateTime) ;
        TBal.PutValue('BSI_DATE2',QExo.FindField('EX_DATEFIN').AsDateTime) ;
      end;
      Ferme(QExo) ;
    // si cumul par mois
    end else if RBS.CumMois then
    begin
      j:=1 ;
      m:=StrToInt(copy(PeriodeMois,5,2)) ;
      a:=StrToInt(copy(PeriodeMois,1,4)) ;
      DatePer:=EncodeDate(a,m,j) ;
      TBal.PutValue('BSI_DATE1',DebutDeMois(DatePer));
      TBal.PutValue('BSI_DATE2',FinDeMois(DatePer));
    end else
    begin
      TBal.PutValue('BSI_DATE1',RBS.DateInf);
      TBal.PutValue('BSI_DATE2',RBS.DateSup);
    end;

    TBal.PutValue('BSI_AXE',RBS.Axe) ;

    TBal.PutValue('BSI_PLAN','N');

    if RBS.Exo='' then
    begin
      if (RBS.CumExo) or (RBS.CumMois) then TBal.PutValue('BSI_EXERCICE',Exercice)
      else TBal.PutValue('BSI_EXERCICE','...' );
    end else TBal.PutValue('BSI_EXERCICE',RBS.Exo);

    if RBS.Devise='' then
    begin
      if RBS.CumDev then TBal.PutValue('BSI_DEVISE',Devise)
      else TBal.PutValue('BSI_DEVISE',V_PGI.DevisePivot);
    end else TBal.PutValue('BSI_DEVISE',RBS.Devise) ;

    if RBS.Etab='' then
    begin
      if RBS.CumEtab then TBal.PutValue('BSI_ETABLISSEMENT',Etablissement)
      else TBal.PutValue('BSI_ETABLISSEMENT','...');
    end else TBal.PutValue('BSI_ETABLISSEMENT',RBS.Etab);

    TBal.PutValue('BSI_SOCIETE',V_PGI.CodeSociete );

    // GCO - 26/06/2006 - Ajotu des balances dynamiques
    TBal.SetString('BSI_TYPEBAL', RBS.TypeBalance);

    // A Ajouter le Type d'écritures
    LibBal := '';
    if RBS.TypeEcr[Reel]  then LibBal := LibBal + 'N';
    if RBS.TypeEcr[Simu]  then LibBal := LibBal + 'S';
    if RBS.TypeEcr[Situ]  then LibBal := LibBal + 'U';
    if RBS.TypeEcr[Previ] then LibBal := LibBal + 'P';
    if RBS.TypeEcr[Revi]  then LibBal := LibBal + 'R';
    if RBS.TypeEcr[Ifrs]  then LibBal := LibBal + 'I';
    TBal.SetString('BSI_TYPEECR', LibBal);
  end;
  Result := TBal;
end;

{$IFNDEF EAGLSERVER}
procedure LookUpBalSit ( E: TControl; TypeCumul, Exo, Etab : string);
var stWhere : string;
begin
  if Exo<> '' then stWhere := 'AND BSI_EXERCICE="'+Exo+'"';
  if Etab <> '' then stWhere := stWhere + 'AND BSI_ETABLISSEMENT="'+Etab+'"';
  if TypeCumul <> '' then stWhere := stWhere+ ' AND BSI_TYPECUM="'+TypeCumul+'"';
  Delete(stWhere,1,4) ;
  LookupList(E,'Balance de situation','CBALSIT','BSI_CODEBAL','BSI_LIBELLE',stWhere,'',True,-1,'',tlLocate);
end;
{$ENDIF}

procedure TBalSit.TriParCompte;
begin
  if (FTypeCumul='G/T') or (FTypeCumul='G/A') then
    LEcr.Detail.Sort('BSE_COMPTE2')
  else LEcr.Detail.Sort('BSE_COMPTE1');
end;

procedure TBalSit.SuppEcritureAZero;
var i,nEcr : integer;
    T : TOB;
begin
  nEcr := LEcr.Detail.Count;
  if nEcr = 0 then exit;
  i := 0;
  T := LEcr.Detail[i];
  while (T <> nil) do
  begin
    if (T.GetValue('BSE_DEBIT')=0) and (T.GetValue('BSE_CREDIT')=0) then
      T.Free
    else Inc (i,1);
    if i < LEcr.Detail.Count then T := LEcr.Detail[i]
    else T := nil;
  end;
  FNbEcriture := LEcr.Detail.Count;
  FModifie := True;
end;

function TBalSit.PresenceCompte(Compte1, Compte2: string): integer;
var i : integer;
    T : TOB;
begin
  Result := -1;
  for i:=0 to LEcr.Detail.Count - 1 do
  begin
    T := LEcr.Detail[i];
    if (T.GetValue ('BSE_COMPTE1')=Compte1) and (T.GetValue ('BSE_COMPTE2')=Compte2) then
    begin
      Result := i;
      break;
    end;
  end;
end;

function TBalSit.SupprimeLigne(Compte1, Compte2: string): boolean;
var i : integer;
    T : TOB;
begin
  Result := False;
  for i:=0 to LEcr.Detail.Count - 1 do
  begin
    T := LEcr.Detail[i];
    if (T.GetValue ('BSE_COMPTE1')=Compte1) and (T.GetValue ('BSE_COMPTE2')=Compte2) then
    begin
      T.Free;
      Result := True;
      break;
    end;
  end;
  FNbEcriture := LEcr.Detail.Count;
  FModifie := True;
end;

end.
