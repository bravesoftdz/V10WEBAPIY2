unit IntegExec;

interface

uses classes, sysutils, dbtables,uTOB,Hctrls,
  // uses Integration
  IntegBase, IntegGen, IntegDef, DefCompte, DefEcriture, DefJournal, DefMois,
  HEnt1, Ent1, SaisUtil, UtilPGI,FE_Main, AGLInit, HMsgBox;

type
  TIntegExec = class (TPersistent)
    private
      fBase : TIntegBase;
      fTOBEcriture : TOB;   // TOB de stockage temporaire des écritures
      fTOBCompte : TOB; // TOB de stockage temporaire des comptes
      fContexte : TIntegContexte;
      procedure InitCommun(TEcr: TOB);
      procedure MajTOBCompte;
      procedure MajTOBEcriture (TDest, TIni : TOB);
      function GetNewNoPiece(PieceSurFolio : boolean;Jal: string;
        DT: TDateTime): integer;
      function ControlePiece(Jal :string; DateComptable : TDateTime): boolean;
    public
      constructor Create (Base : TIntegBase; Contexte : TIntegContexte);
      procedure BuildTOB;
      procedure ToFilePGI (stFileDest : string);
      procedure ToDossier;
    end;

implementation

{ TIntegExec }

procedure TIntegExec.BuildTOB;
var St : string;
  F : TextFile;
  ADefinition : TIntegDefinition;
  T : TOB;
begin
  fTOBCompte := TOB.Create ('', nil , -1);
  fTOBEcriture := TOB.Create ('', nil , -1);
  AssignFile (F,fContexte.FileName);
  Reset (F);
  while not Eof (F) do
  begin
    Readln (F, St);
    ADefinition := fBase.Analyse (St);
    if ADefinition <> nil then
    begin
      if (ADefinition is TIntegDefCompte) then
      begin
        T := TOB.Create ('COMPTE',fTOBCompte,-1);
        TIntegDefCompte(ADefinition).UpdateTOB (T, fContexte);
        fContexte.Ecriture.Compte := T.GetValue('Compte');
      end else
      if (ADefinition is TIntegDefEcriture) then
      begin
        T := TOB.Create ('ECRITURE',fTOBEcriture,-1);
        TIntegDefEcriture(ADefinition).UpdateTOB (T, fContexte);
      end else
      if (ADefinition is TIntegDefJournal) then
      begin
        fContexte.Ecriture.Journal := TIntegDefJournal(ADefinition).GetJournal;
      end else
      if (ADefinition is TIntegDefMois) then
      begin
        TIntegDefMois(ADefinition).GetDate (fContexte.Ecriture);
      end;
    end;
  end;
  CloseFile (F);
  // Mise à jour des natures pour les comptes
  MajTOBCompte;
end;

constructor TIntegExec.Create(Base: TIntegBase; Contexte: TIntegContexte);
begin
  fBase := Base;
  fContexte := Contexte;
end;

procedure TIntegExec.MajTOBEcriture(TDest, TIni: TOB);
begin
  InitCommun (TDest);
  TDest.Parent.PutValue('NumLigne',TDest.Parent.GetValue('NumLigne')+1);
  TDest.PutValue('E_NUMLIGNE',TDest.Parent.GetValue('NumLigne'));
  TDest.PutValue('E_DATEMODIF',NowH);
  TDest.PutValue('E_NUMEROPIECE',TDest.Parent.GetValue('NoPiece'));
  TDest.PutValue('E_ETABLISSEMENT',VH^.EtablisDefaut);
  TDest.PutValue('E_ANA','-');
  TDest.PutValue('E_REFINTERNE',TIni.GetValue ('Piece'));
  TDest.PutValue('E_GENERAL',TIni.GetValue ('General'));
  TDest.PutValue('E_AUXILIAIRE',TIni.GetValue ('Auxiliaire'));
  TDest.PutValue('E_LIBELLE',TIni.GetValue ('Libelle'));
  TDest.PutValue('E_DATECOMPTABLE',TIni.GetValue ('DateComptable'));
  TDest.PutValue('E_DATETAUXDEV',TIni.GetValue ('DateComptable'));
  TDest.PutValue('E_JOURNAL',TIni.GetValue ('Journal'));
  TDest.PutValue('E_DEBIT',TIni.GetValue ('Debit'));
  TDest.PutValue('E_CREDIT',TIni.GetValue ('Credit'));
  TDest.PutValue('E_DEBITDEV',TIni.GetValue ('Debit'));
  TDest.PutValue('E_CREDITDEV',TIni.GetValue ('Credit'));
  TDest.PutValue('E_MODESAISIE','-') ;
  TDest.PutValue('E_EQUILIBRE','-') ;
  TDest.PutValue('E_AVOIRRBT','-') ;
  TDest.PutValue('E_ETAT','0000000000') ;
  TDest.PutValue('E_TYPEMVT','DIV') ;
  if VH^.AttribRibAuto then TDest.PutValue('E_RIB',GetRIBPrincipal(TIni.GetValue ('Auxiliaire')));
{$IFNDEF SPEC302}
  TDest.PutValue('E_PERIODE',GetPeriode(TIni.GetValue ('DateComptable'))) ;
  TDest.PutValue('E_SEMAINE',NumSemaine(TIni.GetValue ('DateComptable'))) ;
{$ENDIF}
  if VH^.TenueEuro then
  begin
    TDest.PutValue('E_DEBITEURO',EuroToPivot(TDest.GetValue('E_DEBIT')));
    TDest.PutValue('E_CREDITEURO',EuroToPivot(TDest.GetValue('E_CREDIT')));
  end
  else
  begin
    TDest.PutValue('E_DEBITEURO',PivotToEuro(TDest.GetValue('E_DEBIT')));
    TDest.PutValue('E_CREDITEURO',PivotToEuro(TDest.GetValue('E_CREDIT')));
  end;
end;

procedure TIntegExec.MajTOBCompte;
var i : integer;
  T : TOB;
  stGeneral , stAuxi : string;
begin
  for i:=0 to fTOBEcriture.Detail.Count - 1 do
  begin
    // Création du compte général dans la TOB compte
    stGeneral := fTOBEcriture.Detail[i].GetValue('General');
    T := fTOBCompte.FindFirst(['General'],[stGeneral],False);
    if T = nil then
    begin
      T:=TOB.Create ('Compte',fTOBCompte,-1);
      if T<> nil then
      begin
        T.AddChampSupValeur('General',stGeneral);
        T.AddChampSupValeur('Libelle',fTOBEcriture.Detail[i].GetValue('Libelle'));
        T.AddChampSupValeur('Auxiliaire','');
      end;
    end;
    // Création du compte auxiliaire dans la TOB compte
    stAuxi := fTOBEcriture.Detail[i].GetValue('Auxiliaire');
    if stAuxi <> '' then
    begin
      T := fTOBCompte.FindFirst(['Auxiliaire'],[stAuxi],False);
      if T = nil then
      begin
        T:=TOB.Create ('Compte',fTOBCompte,-1);
        if T<> nil then
        begin
          T.AddChampSupValeur('General',fTOBEcriture.Detail[i].GetValue('General'));
          T.AddChampSupValeur('Auxiliaire',stAuxi);
          T.AddChampSupValeur('Libelle',fTOBEcriture.Detail[i].GetValue('Libelle'));
        end;
      end;
    end;
  end;
  // Mise à jour de la Nature
  for i:=0 to fTOBCompte.Detail.Count - 1 do
  begin
    T := fTOBCompte.Detail[i];
    // si Auxiliaire vide ==> nature =  nature compte général
    if T.GetValue ('Auxiliaire')='' then
      T.AddChampSupValeur('Nature',IntegGetNatureCompte (fContexte.LCollectif, T.GetValue('General')))
    // sinon nature = nature compte auxiliaire
    else T.AddChampSupValeur('Nature',IntegGetNatureCompte (fContexte.LCollectif, T.GetValue('Auxiliaire')));
  end;
end;

procedure TIntegExec.ToDossier;
var i : integer;
  TEcr,TJal, TPer, T, TDetail : TOB;  // TOB d'écritures rangées par journal et par date
begin
  TEcr := TOB.Create ('', nil, -1);
  // Construction d'un TOB d'écritures "rangées" par journal et par dates
  for i := 0 to fTOBEcriture.Detail.Count - 1 do
  begin
    TDetail := fTOBEcriture.Detail[i];
    // Recherche du journal
    TJal := TEcr.FindFirst(['Journal'],[TDetail.GetValue('Journal')],False);
    if TJal = nil then // Journal introuvable : création de l'enreg.
    begin
      TJal := TOB.Create ('', TEcr,-1);
      TJal.AddChampSupValeur ('Journal', TDetail.GetValue('Journal'));
      TJal.AddChampSupValeur ('PieceSurFolio',PieceSurFolio(TDetail.GetValue('Journal')));
    end;
    // Recherche de la période
    TPer := TJal.FindFirst(['Mois','Annee'],[FormatDateTime('mmmm',TDetail.GetValue('DateComptable')),TDetail.GetValue('Annee')],False);
    if TPer = nil then // Journal introuvable : création de l'enreg.
    begin
      TPer := TOB.Create ('', TJal,-1);
      TPer.AddChampSupValeur ('Mois', FormatDateTime('mmmm',TDetail.GetValue('DateComptable')));
      TPer.AddChampSupValeur  ('Annee', TDetail.GetValue('Annee'));
      TPer.AddChampSupValeur  ('NumLigne', 0);
      if ControlePiece (TJal.GetValue('Journal'),TDetail.GetValue('DateComptable')) then
        TPer.AddChampSupValeur  ('NoPiece', GetNewNoPiece(TJal.GetValue('PieceSurFolio'),TJal.GetValue('Journal'),TDetail.GetValue('DateComptable')))
      else TPer.AddChampSupValeur  ('NoPiece', 0);  // No pièce à 0 , l'écriture n'est pas une écriture de l'exercice en cours.
    end;
    T := TOB.Create ('ECRITURE',TPer,-1);
    MajTOBEcriture (T, TDetail);
  end;
  TheTOB := TEcr;
  AGLLanceFiche('CP','RECUPRESUME','','','');
  if TheTOB <> nil then MessageAlerte(IntToStr(TheTOB.Detail.Count));
  TEcr.Free;
end;

procedure TIntegExec.ToFilePGI(stFileDest: string);
var i : integer;
  St : string;
  F : TextFile;
  T : TOB;
begin
  AssignFile (F, stFileDest);
  Rewrite(F);
  // Enregistrement des comptes
  for i := 0 to fTOBCompte.Detail.Count  - 1 do
  begin
    T := fTOBCompte.Detail[i];
    // Si Nature='FOU' ou nature = 'CLI' ==> Auxiliaire
    if ((T.GetValue('Nature')='FOU') or (T.GetValue('Nature')='CLI')) then
      St := Format(FMT_AUXILIAIRE,['###','CAU',T.GetValue('Auxiliaire'),
                      T.GetValue('Libelle'),T.GetValue('Nature'),'X',
                      T.GetValue('General'),'','FRA','CHQ'])
    // Sinon général (Nature = 'DIV', 'COC' ou 'COF'
    else St := Format(FMT_GENERAL,['###','CGE',T.GetValue('General'),
                      T.GetValue('Libelle'),T.GetValue('Nature'),'-','-','-',
                      '-','-','-','-']);
    Writeln(F,St);
  end;
  // Enregistrement des écritures
  for i := 0 to fTOBEcriture.Detail.Count  - 1 do
  begin
    St := Format(FMT_ECRITURE,[fTOBEcriture.Detail[i].GetValue('Journal'), fTOBEcriture.Detail[i].GetValue('DateComptable'),'OD',fTOBEcriture.Detail[i].GetValue('General'),
      '',fTOBEcriture.Detail[i].GetValue('Auxiliaire'),'',fTOBEcriture.Detail[i].GetValue('Libelle'),'CHQ',fTOBEcriture.Detail[i].GetValue('DateEcheance'),
      fTOBEcriture.Detail[i].GetValue('Sens'),fTOBEcriture.Detail[i].GetValue('Montant'),'N',
      fTOBEcriture.Detail[i].GetValue('Piece')]);
    Writeln(F,St);
  end;
  CloseFile(F);
end;

procedure TIntegExec.InitCommun(TEcr : TOB);
begin
  TEcr.PutValue('E_EXERCICE',VH^.EnCours.Code);
  TEcr.PutValue('E_NATUREPIECE','OD');
  TEcr.PutValue('E_QUALIFPIECE','N');
  TEcr.PutValue('E_DATECREATION',Date);
  TEcr.PutValue('E_DEVISE',V_PGI.DevisePivot);
  TEcr.PutValue('E_TAUXDEV',1);
  TEcr.PutValue('E_COTATION',1);
  TEcr.PutValue('E_SAISIEEURO','-');
  TEcr.PutValue('E_LETTRAGEEURO','-');
  TEcr.PutValue('E_CREERPAR','GEN');
  TEcr.PutValue('E_QUALIFORIGINE','IMO');
  TEcr.PutValue('E_ECRANOUVEAU','N');
  TEcr.PutValue('E_ETATLETTRAGE','RI');
  TEcr.PutValue('E_UTILISATEUR',V_PGI.User);
  TEcr.PutValue('E_CONFIDENTIEL','0');
  TEcr.PutValue('E_SOCIETE',V_PGI.CodeSociete);
end;

function TIntegExec.GetNewNoPiece(PieceSurFolio : boolean;Jal : string; DT : TDateTime) : integer;
var NoPiece : integer;
    Year, Month, Day : Word;
    stWhere : string;
    Q : TQuery;
begin
  if (not PieceSurFolio) then
  begin
    NoPiece := GetNewNumJal(Jal,True, DT);
  end else
  begin
    DecodeDate(DT, Year, Month, Day) ;
    stWhere := 'WHERE E_EXERCICE="'+VH^.Encours.Code+'" AND E_JOURNAL="'+Jal+'"';
    stWhere := stWhere +' AND E_DATECOMPTABLE>="'+USDateTime(EncodeDate(Year, Month, 1))+'"' ;
    stWhere := stWhere +' AND E_DATECOMPTABLE<="'+USDateTime(EncodeDate(Year, Month, DaysPerMonth(Year, Month)))+'"' ;
    Q := OpenSQL ('SELECT MAX(E_NUMEROPIECE) FROM ECRITURE '+stWhere,True);
    if not Q.Eof then NoPiece := (Q.Fields[0].AsInteger+1)
    else NoPiece := 1; // Premier folio
    Ferme (Q);
  end;
  result := NoPiece;
end;

function TIntegExec.ControlePiece(Jal :string; DateComptable : TDateTime): boolean;
begin
  Result := ((Presence ('JOURNAL','J_JOURNAL',Jal)) and
          (DateComptable<=VH^.Encours.Fin) and
          (DateComptable>=VH^.Encours.Deb));
end;

end.
