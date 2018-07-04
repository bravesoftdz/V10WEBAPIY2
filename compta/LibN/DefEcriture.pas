unit DefEcriture;

interface

uses SysUtils, uTOB, HCtrls, HEnt1, IntegDef , IntegGen;

type
  TIntegDefEcriture = class (TIntegDefinition)
    private
      DateComptable : TDateTime;
    public
      Jour : TIntegZoneDef;
      Mois : TIntegZoneDef;
      Annee : TIntegZoneDef;
      Compte : TIntegZoneDef;
      Libelle : TIntegZoneDef;
      Stat : TIntegZoneDef;
      Piece : TIntegZoneDef;
      Debit : TIntegZoneDef;
      Credit : TIntegZoneDef;
      Lettrage : TIntegZoneDef;
      DebitN1 : TIntegZoneDef;
      CreditN1 : TIntegZoneDef;
      Quantite : TIntegZoneDef;
      Reglement : TIntegZoneDef;
      JourEche : TIntegZoneDef;
      MoisEche : TIntegZoneDef;
      AnneeEche : TIntegZoneDef;
      Journal  : TIntegZoneDef;
      LibCompte : TIntegZoneDef;
      constructor Create; 
      procedure Charge (var F : TextFile);
      procedure Enregistre (var F : TextFile);
      procedure MajValeur (St : string);
      procedure UpdateTOB(T: TOB; Ctx : TIntegContexte);
      procedure UpdateGrille ( G : THGrid);
  end;

function FormateDate (DT : TDateTime) : string;

implementation

{ TIntegDefEcriture }

procedure TIntegDefEcriture.Charge(var F: TextFile);
var St, stEntete : string;
begin
  repeat
    Readln (F, St);
    stEntete := IntegExtractEntete (St);
    if stEntete = EX_EPOSDJ then Jour.Pos := IntegExtractValeur(St)+1
    else if stEntete = EX_EPOSDM then Mois.Pos := IntegExtractValeur(St)+1
    else if stEntete = EX_EPOSDA then Annee.Pos := IntegExtractValeur(St)+1
    else if stEntete = EX_EPOSNC then Compte.Pos := IntegExtractValeur(St)+1
    else if stEntete = EX_ETAILLENC then Compte.Lg := IntegExtractValeur(St)
    else if stEntete = EX_EPOSLC then Libelle.Pos := IntegExtractValeur(St)+1
    else if stEntete = EX_ETAILLELC then Libelle.Lg := IntegExtractValeur(St)
    else if stEntete = EX_EPOSS then Stat.Pos := IntegExtractValeur(St)+1
    else if stEntete = EX_ETAILLES then Stat.Lg := IntegExtractValeur(St)
    else if stEntete = EX_EPOSP then Piece.Pos := IntegExtractValeur(St)+1
    else if stEntete = EX_ETAILLEP then Piece.Lg := IntegExtractValeur(St)
    else if stEntete = EX_EPOSDB then Debit.Pos := IntegExtractValeur(St)+1
    else if stEntete = EX_ETAILLEDB then Debit.Lg := IntegExtractValeur(St)
    else if stEntete = EX_EPOSCR then Credit.Pos := IntegExtractValeur(St)+1
    else if stEntete = EX_ETAILLECR then Credit.Lg := IntegExtractValeur(St)
    else if stEntete = EX_EPOSA then Lettrage.Pos := IntegExtractValeur(St)+1
    else if stEntete = EX_ETAILLEA then Lettrage.Lg := IntegExtractValeur(St)
    else if stEntete = EX_EPOSD1 then DebitN1.Pos := IntegExtractValeur(St)+1
    else if stEntete = EX_ETAILLED1 then DebitN1.Lg := IntegExtractValeur(St)
    else if stEntete = EX_EPOSC1 then CreditN1.Pos := IntegExtractValeur(St)+1
    else if stEntete = EX_ETAILLEC1 then CreditN1.Lg := IntegExtractValeur(St)
    else if stEntete = EX_EPOSQ then Quantite.Pos := IntegExtractValeur(St)+1
    else if stEntete = EX_ETAILLEQ then Quantite.Lg := IntegExtractValeur(St)
    else if stEntete = EX_EPOSRG then Reglement.Pos := IntegExtractValeur(St)+1
    else if stEntete = EX_ETAILLERG then Reglement.Lg := IntegExtractValeur(St)
    else if stEntete = EX_EPOSEJ then JourEche.Pos := IntegExtractValeur(St)+1
    else if stEntete = EX_EPOSEM then MoisEche.Pos := IntegExtractValeur(St)+1
    else if stEntete = EX_EPOSEA then AnneeEche.Pos := IntegExtractValeur(St)+1
    else if stEntete = EX_EPOSJ then Journal.Pos := IntegExtractValeur(St)+1
    else if stEntete = EX_ETAILLEJ then Journal.Lg := IntegExtractValeur(St)
    else if stEntete = EX_EPOSLJ then LibCompte.Pos := IntegExtractValeur(St)+1
  until ((St='') or (stEntete=EX_ETAILLELJ));
  if stEntete = EX_ETAILLELJ then LibCompte.Lg := IntegExtractValeur(St);
end;

constructor TIntegDefEcriture.Create;
begin
  inherited;
  SetType (ilEcriture);
  DateComptable := iDate1900;
  Jour.Val := 0; Mois.Val := 0; Annee.Val := 0;
  Compte.Val := ''; Libelle.Val := ''; Stat.Val := ''; Piece.Val := '';
  Debit.Val := 0; Credit.Val := 0; DebitN1.Val := 0; CreditN1.Val := 0;
  Quantite.Val := 0; Reglement.Val := ''; JourEche.Val := 0; MoisEche.Val := 0; AnneeEche.Val := 0;
  Journal.Val := ''; LibCompte.Val := '';
end;

procedure TIntegDefEcriture.Enregistre(var F: TextFile);
begin
  Writeln (F, EX_EPOSDJ+'  '+Format('%.04d',[Jour.Pos-1]));
  Writeln (F, EX_EPOSDM+'  '+Format('%.04d',[Mois.Pos-1]));
  Writeln (F, EX_EPOSDA+'  '+Format('%.04d',[Annee.Pos-1]));
  Writeln (F, EX_EPOSNC+'  '+Format('%.04d',[Compte.Pos-1]));
  Writeln (F, EX_ETAILLENC+'  '+Format('%.04d',[Compte.Lg]));
  Writeln (F, EX_EPOSLC+'  '+Format('%.04d',[Libelle.Pos-1]));
  Writeln (F, EX_ETAILLELC+'  '+Format('%.04d',[Libelle.Lg]));
  Writeln (F, EX_EPOSS+'  '+Format('%.04d',[Stat.Pos-1]));
  Writeln (F, EX_ETAILLES+'  '+Format('%.04d',[Stat.Lg]));
  Writeln (F, EX_EPOSP+'  '+Format('%.04d',[Piece.Pos-1]));
  Writeln (F, EX_ETAILLEP+'  '+Format('%.04d',[Piece.Lg]));
  Writeln (F, EX_EPOSDB+'  '+Format('%.04d',[Debit.Pos-1]));
  Writeln (F, EX_ETAILLEDB+'  '+Format('%.04d',[Debit.Lg]));
  Writeln (F, EX_EPOSCR+'  '+Format('%.04d',[Credit.Pos-1]));
  Writeln (F, EX_ETAILLECR+'  '+Format('%.04d',[Credit.Lg]));
  Writeln (F, EX_EPOSA+'  '+Format('%.04d',[Lettrage.Pos-1]));
  Writeln (F, EX_ETAILLEA+'  '+Format('%.04d',[Lettrage.Lg]));
  Writeln (F, EX_EPOSD1+'  '+Format('%.04d',[DebitN1.Pos-1]));
  Writeln (F, EX_ETAILLED1+'  '+Format('%.04d',[DebitN1.Lg]));
  Writeln (F, EX_EPOSC1+'  '+Format('%.04d',[CreditN1.Pos-1]));
  Writeln (F, EX_ETAILLEC1+'  '+Format('%.04d',[CreditN1.Lg]));
  Writeln (F, EX_EPOSQ+'  '+Format('%.04d',[Quantite.Pos-1]));
  Writeln (F, EX_ETAILLEQ+'  '+Format('%.04d',[Quantite.Lg]));
  Writeln (F, EX_EPOSRG+'  '+Format('%.04d',[Reglement.Pos-1]));
  Writeln (F, EX_ETAILLERG+'  '+Format('%.04d',[Reglement.Lg]));
  Writeln (F, EX_EPOSEJ+'  '+Format('%.04d',[JourEche.Pos-1]));
  Writeln (F, EX_EPOSEM+'  '+Format('%.04d',[MoisEche.Pos-1]));
  Writeln (F, EX_EPOSEA+'  '+Format('%.04d',[AnneeEche.Pos-1]));
  Writeln (F, EX_EPOSJ+'  '+Format('%.04d',[Journal.Pos-1]));
  Writeln (F, EX_ETAILLEJ+'  '+Format('%.04d',[Journal.Lg]));
  Writeln (F, EX_EPOSLJ+'  '+Format('%.04d',[LibCompte.Pos-1]));
  Writeln (F, EX_ETAILLELJ+'  '+Format('%.04d',[LibCompte.Lg]));
end;

procedure TIntegDefEcriture.MajValeur(St: string);
var Year : Word;
begin
      Jour.Val := Valeur(Copy (St,Jour.Pos,2));
      Mois.Val := Valeur(Copy (St,Mois.Pos,2));
      Annee.Val := Valeur(Copy (St,Annee.Pos,2));
      if IsValidDate(Format('%02d/%02d/%02d',[integer(Jour.Val),integer(Mois.Val),integer(Annee.Val)])) then
      begin
        if Annee.Val < 70 then Year := 2000+Annee.Val
        else if Annee.Val<100 then Year := 1900+Annee.Val
        else Year := Annee.Val;
        DateComptable := EncodeDate(Year, Mois.Val, Jour.Val);
      end;
      Compte.Val := Copy (St,Compte.Pos,Compte.Lg);
      Libelle.Val := Copy (St,Libelle.Pos,Libelle.Lg);
      Stat.Val := Copy (St,Stat.Pos,Stat.Lg);
      Piece.Val := Copy (St,Piece.Pos,Piece.Lg);
      Debit.Val := Valeur(Copy (St,Debit.Pos,Debit.Lg));
      Credit.Val := Valeur(Copy (St,Credit.Pos,Credit.Lg));
      Lettrage.Val := Copy (St,Lettrage.Pos,Lettrage.Lg);
      DebitN1.Val := Valeur(Copy (St,DebitN1.Pos,DebitN1.Lg));
      CreditN1.Val := Valeur(Copy (St,CreditN1.Pos,CreditN1.Lg));
      Quantite.Val := Valeur(Copy (St,Quantite.Pos,Quantite.Lg));
      Reglement.Val := Copy (St,Reglement.Pos,Reglement.Lg);
      JourEche.Val := Valeur(Copy (St,JourEche.Pos,2));
      MoisEche.Val := Valeur(Copy (St,MoisEche.Pos,2));
      AnneeEche.Val := Valeur(Copy (St,AnneeEche.Pos,2));
      Journal.Val := Copy (St,Journal.Pos,Journal.Lg);
      LibCompte.Val := Copy (St,LibCompte.Pos,LibCompte.Lg);
end;

procedure TIntegDefEcriture.UpdateGrille(G: THGrid);
begin                          
  G.RowCount := 9;
  G.Cells[0,0] := 'Compte'; G.Cells[1,0] := Compte.Val;
  G.Cells[0,1] := 'Journal'; G.Cells[1,1] := Journal.Val;
  G.Cells[0,2] := 'Date'; G.Cells[1,2] := '  /  /    ';
  G.Cells[0,3] := 'Libellé'; G.Cells[1,3] := Libelle.Val;
  G.Cells[0,4] := 'Pièce'; G.Cells[1,4] := Piece.Val;
  G.Cells[0,5] := 'Débit'; G.Cells[1,5] := StrfMontant(Debit.Val,17,2,'',True);
  G.Cells[0,6] := 'Crédit'; G.Cells[1,6] := StrfMontant(Credit.Val,17,2,'',True);
  G.Cells[0,7] := 'Lettrage'; G.Cells[1,7] := Lettrage.Val;
  G.Cells[0,8] := 'Echeance'; G.Cells[1,8] := '  /  /    ';
end;

procedure TIntegDefEcriture.UpdateTOB(T: TOB; Ctx : TIntegContexte);
var stCompte, stJournal, stDate, stAuxi : string;
  jj,mm,aa : Word;
begin
  stAuxi := '';
  if Compte.Val = '' then stCompte := Ctx.Ecriture.Compte else stCompte := Compte.Val;
  if Journal.Val = '' then stJournal := Ctx.Ecriture.Journal else stJournal := Journal.Val;

  if DateComptable=iDate1900 then
  begin
    if Jour.Val = 0 then
      if Ctx.Ecriture.Jour > 0 then jj := Ctx.Ecriture.Jour else jj := 1
    else jj := Jour.Val;
    if Mois.Val = 0 then
      if Ctx.Ecriture.Mois > 0 then mm := Ctx.Ecriture.Mois else mm := 1
    else mm := Mois.Val;
    if Annee.Val = 0 then
      if Ctx.Ecriture.Annee > 0 then
      begin
        if Ctx.Ecriture.Annee > 70 then aa := 1900+Ctx.Ecriture.Annee
        else aa := 2000+Ctx.Ecriture.Annee;
      end else aa := 2001
    else aa := Annee.Val;
    if IsValidDate(Format('%02d/%02d/%02d',[jj,mm,aa])) then stDate := FormateDate(EnCodeDate(aa,mm,jj))
    else stDate := FormateDate(EnCodeDate(2001,01,01));
  end else stDate := FormateDate(DateComptable);
  DecodeDate(DateComptable,aa,mm,jj);
  T.AddChampSupValeur ('Mois',mm);
  T.AddChampSupValeur ('Annee',aa);
  T.AddChampSupValeur('DateComptable',EncodeDate(aa,mm,jj));

  if ((JourEche.Val=0) or (MoisEche.Val=0) or (AnneeEche.Val=0)) then
    T.AddChampSupValeur('DateEcheance','        ')
  else T.AddChampSupValeur('DateEcheance',Format('%02d%02d%04d',[integer(JourEche.Val),integer(MoisEche.Val),integer(AnneeEche.Val)]));

  IntegGetCompteEquiv (Ctx ,stCompte, stAuxi);
  T.AddChampSupValeur('General',stCompte);
  T.AddChampSupValeur('Auxiliaire',stAuxi);
  T.AddChampSupValeur('Libelle',Libelle.Val);
  T.AddChampSupValeur('Stat',Stat.Val);
  T.AddChampSupValeur('Piece',Piece.Val);
  if Debit.Val > 0 then
  begin
    T.AddChampSupValeur('Sens','D');
    T.AddChampSupValeur('Montant',Debit.Val);
    T.AddChampSupValeur('Debit',Debit.Val);
    T.AddChampSupValeur('Credit',0);
  end else
  begin
    T.AddChampSupValeur('Sens','C');
    T.AddChampSupValeur('Montant',Credit.Val);
    T.AddChampSupValeur('Debit',0);    
    T.AddChampSupValeur('Credit',Credit.Val);
  end;
  T.AddChampSupValeur('Lettrage',Lettrage.Val);
  // T.AddChampSupValeur('Quantite',Quantite.Val);
  // T.AddChampSupValeur('Reglement',Reglement.Val);
  T.AddChampSupValeur('Journal',stJournal);
  // T.AddChampSupValeur('LibCompte',LibCompte.Val);
end;

function FormateDate (DT : TDateTime) : string;
var Y, M, D : Word;
begin
  if DT = iDate1900 then Result := '' else
  begin
    DecodeDate(DT,Y,M,D);
    Result := Format('%.02d%.02d%.04d',[D,M,Y]);
  end;
end;

end.


