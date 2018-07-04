unit P5UtilCongesPayes;

interface
uses
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  SysUtils, HCtrls, HEnt1, ed_tools, Hdebug,
{$IFNDEF EAGLCLIENT}
  Db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
(*  Mise en commentaire des uses inutiles
  Fiche,fichlist,
{$ELSE}
  eFiche, eFichList,
{$ENDIF}
*)

  Controls, HMsgBox, UTOB;

const
  DCP: integer = 2; // Nbre de décimales pour le calcul des congés payés

function  PGCalculValoCP(LCTob_GblCp, LCTobEtab, LCTobSAL : tob; LcNbJOuvres, LcNbJOuvrables: double; LcDatedeb, Lcdatefin: tdatetime; var LcNext: integer; LcTypeCp: string): boolean;
Function  PGCalculCompteurCp(LCTob_GblCp : Tob): Tob;
Function  PGRecupTob_pris(Lc_tobCP : Tob; DCP : TDateTime; LcTypeCp : String; ChgePar : Boolean = False) : Tob;
Function  PGRecupTob_prisPayes(Lc_tobCP : Tob;  DCP : TDateTime; LcTypeCp : String; ChgePar : Boolean = False) : Tob;
Function  PGRecupTob_Acquis(Lc_tobCP : Tob; DCP : TDateTime; ChgePar : Boolean = False) : Tob;
Function  PGRecupTob_AjuAjp(Lc_tobCP : Tob; DCP : TDateTime; ChgePar : Boolean = False) : Tob;
procedure PgReintegreTobMere(LcTob,LcTob_GblCp : Tob);

Function  PGExistSoldeCloture(Lc_tobCP: tob; DCP : TDateTime) : Boolean;

procedure PgGenereSolde(Tob_Sal, Tob_etab, LCTob_GblCp : tob; DateF, DateD: tdatetime; var next: integer; Action: string);

{ Méthode de remplacement des variables globales }
Procedure  PgGetJHCPPris( LCTob_GblCp : Tob; DCP : TDateTime; Var Nbj, Nbh : double) ;
Procedure  PgGetJHCPPrisPoses( LCTob_GblCp : Tob; DCP : TDateTime; Var Nbj, Nbh : double) ;
Procedure  PgGetMTCPPris( LCTob_GblCp : Tob; DCP : TDateTime; Var Abs, Ind : double) ;
Procedure  PgGetJHCPSolde( LCTob_GblCp : Tob; DCP : TDateTime; Var Nbj, Nbh : double) ;
Procedure  PgGetMTCPSolde( LCTob_GblCp : Tob; DCP : TDateTime; Var Abs, Ind : double) ;

Function   PgGetJCPArr ( LCTob_GblCp : Tob; DCP : TDateTime) : double ;



implementation
uses pgcongespayes,pgcommun,EntPaie,P5Util;

Function PGCalculCompteurCp(LCTob_GblCp : Tob ): Tob;
Var


T1 : Tob;
Mt1, Mt2, MtRel  : Double;
Begin
  Result := nil;
{  St := 'SELECT * FROM ABSENCESALARIE WHERE PCN_TYPEMVT="CPA" '+
        'AND PCN_SALARIE="'+Sal+'" AND PCN_PERIODECP<2';
  Q := OpenSQL(St,True);
  If not Q.eof then
    Begin
    Tob_Cp := TOB.Create('les Cp',nil,-1);
    Tob_Cp.LoadDetailDB('','','',Q,False);
    End;
  Ferme(Q);}

  If Assigned(LCTob_GblCp) then
     Begin
     Result := TOB.Create('les compteurs Cp',nil,-1);
     { Traitement des Reliquats }
     Mt1 := LCTob_GblCp.Somme('PCN_JOURS',['PCN_TYPECONGE','PCN_PERIODECP','PCN_SENSABS'],['REL','1','+'],False);
     Mt1 := Mt1 - LCTob_GblCp.Somme('PCN_JOURS',['PCN_TYPECONGE','PCN_PERIODECP','PCN_SENSABS'],['REL','1','-'],False);
     Mt1 := Mt1 + LCTob_GblCp.Somme('PCN_JOURS',['PCN_TYPECONGE','PCN_PERIODECP','PCN_SENSABS','PCN_TYPEIMPUTE'],['AJU','1','+','REL'],False);
     Mt1 := Mt1 - LCTob_GblCp.Somme('PCN_JOURS',['PCN_TYPECONGE','PCN_PERIODECP','PCN_SENSABS','PCN_TYPEIMPUTE'],['AJU','1','-','REL'],False);
     Mt2 := LCTob_GblCp.Somme('PCN_BASE',['PCN_TYPECONGE','PCN_PERIODECP','PCN_SENSABS'],['REL','1','+'],False);
     Mt2 := Mt2 - LCTob_GblCp.Somme('PCN_BASE',['PCN_TYPECONGE','PCN_PERIODECP','PCN_SENSABS'],['REL','1','-'],False);
     Mt2 := Mt2 + LCTob_GblCp.Somme('PCN_BASE',['PCN_TYPECONGE','PCN_PERIODECP','PCN_SENSABS','PCN_TYPEIMPUTE'],['AJU','1','+','REL'],False);
     Mt2 := Mt2 - LCTob_GblCp.Somme('PCN_BASE',['PCN_TYPECONGE','PCN_PERIODECP','PCN_SENSABS','PCN_TYPEIMPUTE'],['AJU','1','-','REL'],False);
     Result.AddChampSupValeur('RELIQUAT',Mt1);
     Result.AddChampSupValeur('BASERELIQUAT',Mt2);
     { Traitement des acquis N-1 }
     Mt1 := LCTob_GblCp.Somme('PCN_JOURS',['PCN_TYPECONGE','PCN_PERIODECP'],['ACQ','1'],False);
     Mt1 := Mt1 + LCTob_GblCp.Somme('PCN_JOURS',['PCN_TYPECONGE','PCN_PERIODECP'],['ACS','1'],False);
     Mt1 := Mt1 + LCTob_GblCp.Somme('PCN_JOURS',['PCN_TYPECONGE','PCN_PERIODECP'],['ACA','1'],False);
     Mt1 := Mt1 + LCTob_GblCp.Somme('PCN_JOURS',['PCN_TYPECONGE','PCN_PERIODECP'],['REP','1'],False);
     Mt1 := Mt1 + LCTob_GblCp.Somme('PCN_JOURS',['PCN_TYPECONGE','PCN_PERIODECP'],['ARR','1'],False);
     Mt1 := Mt1 + LCTob_GblCp.Somme('PCN_JOURS',['PCN_TYPECONGE','PCN_PERIODECP','PCN_SENSABS','PCN_TYPEIMPUTE'],['AJU','1','+','ACQ'],False);
     Mt1 := Mt1 - LCTob_GblCp.Somme('PCN_JOURS',['PCN_TYPECONGE','PCN_PERIODECP','PCN_SENSABS','PCN_TYPEIMPUTE'],['AJU','1','-','ACQ'],False);
     Mt2 := LCTob_GblCp.Somme('PCN_BASE',['PCN_TYPECONGE','PCN_PERIODECP'],['ACQ','1'],False);
     Mt2 := Mt2 + LCTob_GblCp.Somme('PCN_BASE',['PCN_TYPECONGE','PCN_PERIODECP'],['ACS','1'],False);
     Mt2 := Mt2 + LCTob_GblCp.Somme('PCN_BASE',['PCN_TYPECONGE','PCN_PERIODECP'],['ACA','1'],False);
     Mt2 := Mt2 + LCTob_GblCp.Somme('PCN_BASE',['PCN_TYPECONGE','PCN_PERIODECP'],['REP','1'],False);
     Mt2 := Mt2 + LCTob_GblCp.Somme('PCN_BASE',['PCN_TYPECONGE','PCN_PERIODECP'],['ARR','1'],False);
     Mt2 := Mt2 + LCTob_GblCp.Somme('PCN_BASE',['PCN_TYPECONGE','PCN_PERIODECP','PCN_SENSABS','PCN_TYPEIMPUTE'],['AJU','1','+','ACQ'],False);
     Mt2 := Mt2 - LCTob_GblCp.Somme('PCN_BASE',['PCN_TYPECONGE','PCN_PERIODECP','PCN_SENSABS','PCN_TYPEIMPUTE'],['AJU','1','-','ACQ'],False);
     Result.AddChampSupValeur('ACQUISN1',Mt1);
     Result.AddChampSupValeur('BASEACQUISN1',Mt2);
     { Traitement des acquis N }
     Mt1 := LCTob_GblCp.Somme('PCN_JOURS',['PCN_TYPECONGE','PCN_PERIODECP'],['ACQ','0'],False);
     Mt1 := Mt1 + LCTob_GblCp.Somme('PCN_JOURS',['PCN_TYPECONGE','PCN_PERIODECP'],['ACS','0'],False);
     Mt1 := Mt1 + LCTob_GblCp.Somme('PCN_JOURS',['PCN_TYPECONGE','PCN_PERIODECP'],['ACA','0'],False);
     Mt1 := Mt1 + LCTob_GblCp.Somme('PCN_JOURS',['PCN_TYPECONGE','PCN_PERIODECP'],['REP','0'],False);
     Mt1 := Mt1 + LCTob_GblCp.Somme('PCN_JOURS',['PCN_TYPECONGE','PCN_PERIODECP'],['ARR','0'],False);
     Mt1 := Mt1 + LCTob_GblCp.Somme('PCN_JOURS',['PCN_TYPECONGE','PCN_PERIODECP','PCN_SENSABS','PCN_TYPEIMPUTE'],['AJU','0','+','ACQ'],False);
     Mt1 := Mt1 - LCTob_GblCp.Somme('PCN_JOURS',['PCN_TYPECONGE','PCN_PERIODECP','PCN_SENSABS','PCN_TYPEIMPUTE'],['AJU','0','-','ACQ'],False);
     Mt2 := LCTob_GblCp.Somme('PCN_BASE',['PCN_TYPECONGE','PCN_PERIODECP'],['ACQ','0'],False);
     Mt2 := Mt2 + LCTob_GblCp.Somme('PCN_BASE',['PCN_TYPECONGE','PCN_PERIODECP'],['ACS','0'],False);
     Mt2 := Mt2 + LCTob_GblCp.Somme('PCN_BASE',['PCN_TYPECONGE','PCN_PERIODECP'],['ACA','0'],False);
     Mt2 := Mt2 + LCTob_GblCp.Somme('PCN_BASE',['PCN_TYPECONGE','PCN_PERIODECP'],['REP','0'],False);
     Mt2 := Mt2 + LCTob_GblCp.Somme('PCN_BASE',['PCN_TYPECONGE','PCN_PERIODECP'],['ARR','0'],False);
     Mt2 := Mt2 + LCTob_GblCp.Somme('PCN_BASE',['PCN_TYPECONGE','PCN_PERIODECP','PCN_SENSABS','PCN_TYPEIMPUTE'],['AJU','0','+','ACQ'],False);
     Mt2 := Mt2 - LCTob_GblCp.Somme('PCN_BASE',['PCN_TYPECONGE','PCN_PERIODECP','PCN_SENSABS','PCN_TYPEIMPUTE'],['AJU','0','-','ACQ'],False);
     Result.AddChampSupValeur('ACQUISN',Mt1);
     Result.AddChampSupValeur('BASEACQUISN',Mt2);
     { Traitement des pri }
     Mt1 := 0; Mt2 := 0;  MtRel := 0;
     T1 := LCTob_GblCp.FindFirst(['PCN_TYPECONGE','PCN_PERIODECP'],['PRI','1'],False);
     While Assigned(T1) do
       Begin
       if ((T1.GetValue('PCN_MVTDUPLIQUE')='X') OR (T1.GetValue('PCN_CODERGRPT')<>'-1'))
       AND (T1.GetValue('PCN_CODETAPE')<>'...' )  AND  (T1.GetValue('PCN_ETATPOSTPAIE')<>'NAN') then
          Begin
          if (POS('Reliquat',T1.GetValue('PCN_PERIODEPAIE'))> 0) then
              MtRel := MtRel + T1.GetValue('PCN_JOURS')
          else
            Begin
            if (T1.GetValue('PCN_PERIODECP')='1') then Mt1 := Mt1 + T1.GetValue('PCN_JOURS')
            else                                       Mt2 := Mt2 + T1.GetValue('PCN_JOURS');
            End;
          End;
       T1 := LCTob_GblCp.FindNext(['PCN_TYPECONGE','PCN_PERIODECP'],['PRI','1'],False);
       End;
     T1 := LCTob_GblCp.FindFirst(['PCN_TYPECONGE','PCN_PERIODECP'],['SLD','1'],False);
     While Assigned(T1) do
       Begin
       if ((T1.GetValue('PCN_MVTDUPLIQUE')='X') OR (T1.GetValue('PCN_CODERGRPT')<>'-1')) AND (T1.GetValue('PCN_CODETAPE')<>'...' ) then
          Begin
          if POS('Reliquat',T1.GetValue('PCN_PERIODEPAIE'))> 0 then
              MtRel := MtRel + T1.GetValue('PCN_JOURS')
          else
            Begin
            if (T1.GetValue('PCN_PERIODECP')='1') then Mt1 := Mt1 + T1.GetValue('PCN_JOURS')
            else                                       Mt2 := Mt2 + T1.GetValue('PCN_JOURS');
            End;
          End;
       T1 := LCTob_GblCp.FindNext(['PCN_TYPECONGE','PCN_PERIODECP'],['SLD','1'],False);
       End;
     { Traitement des pri sur reliquat }
     Result.AddChampSupValeur('PRIRELIQUAT',MtRel);
     { Traitement des pri N-1 }
     Mt1 := Mt1 + LCTob_GblCp.Somme('PCN_JOURS',['PCN_TYPECONGE','PCN_PERIODECP'],['CPA','1'],False);
     Mt1 := Mt1 - LCTob_GblCp.Somme('PCN_JOURS',['PCN_TYPECONGE','PCN_PERIODECP','PCN_SENSABS'],['AJP','1','+'],False);
     Mt1 := Mt1 + LCTob_GblCp.Somme('PCN_JOURS',['PCN_TYPECONGE','PCN_PERIODECP','PCN_SENSABS'],['AJP','1','-'],False);
     Result.AddChampSupValeur('PRIN1',Mt1);
     { Traitement des pri N }
     Mt2 := Mt2 + LCTob_GblCp.Somme('PCN_JOURS',['PCN_TYPECONGE','PCN_PERIODECP'],['CPA','0'],False);
     Mt2 := Mt2 - LCTob_GblCp.Somme('PCN_JOURS',['PCN_TYPECONGE','PCN_PERIODECP','PCN_SENSABS'],['AJP','0','+'],False);
     Mt2 := Mt2 + LCTob_GblCp.Somme('PCN_JOURS',['PCN_TYPECONGE','PCN_PERIODECP','PCN_SENSABS'],['AJP','0','-'],False);
     Result.AddChampSupValeur('PRIN',Mt2);
     End;
End;

Function PGRecupTob_pris(Lc_tobCP : Tob;  DCP : TDateTime; LcTypeCp : String; ChgePar : Boolean = False) : Tob;
Var
T1, T2 : Tob;
Begin
  Result := Tob.create('Les pris',nil,-1);
  T1 :=  Lc_tobCP.FindFirst([''],[''],False);
  While Assigned(T1) do
    Begin
    If (((T1.getValue('PCN_TYPECONGE') = 'PRI') AND (T1.getValue('PCN_VALIDRESP') = 'VAL'))
    OR (T1.getValue('PCN_TYPECONGE') = 'SLD'))
    AND (T1.getValue('PCN_DATEVALIDITE') <= DCP) AND (T1.getValue('PCN_DATEFIN') <= 10) then
      Begin
      If (LcTypeCp<>'') and (T1.getValue('PCN_TYPECONGE') <> LcTypeCp )then
        Begin
        T1 :=  Lc_tobCP.FindNext([''],[''],False);
        continue;
        End;
      if ChgePar then
        T1.ChangeParent(Result, -1)
      else
        Begin
        T2 := Tob.create('Les pris',Result,-1);
        T2.Dupliquer(T1,True,True,False);
        End;
      End;
    T1 :=  Lc_tobCP.FindNext([''],[''],False);
    End;
  If Result.detail.count = 0 then FreeAndNil(Result)
  else
     Result.detail.Sort('PCN_DATEVALIDITE');
End;


Function PGRecupTob_prisPayes(Lc_tobCP : Tob;  DCP : TDateTime; LcTypeCp : String; ChgePar : Boolean = False) : Tob;
Var
T1, T2 : Tob;
Begin
  Result := Tob.create('Les pris payés',nil,-1);
  T1 :=  Lc_tobCP.FindFirst([''],[''],False);
  While Assigned(T1) do
    Begin
    If (((T1.getValue('PCN_TYPECONGE') = 'PRI') AND (T1.getValue('PCN_VALIDRESP') = 'VAL'))
    OR (T1.getValue('PCN_TYPECONGE') = 'SLD'))
    AND (T1.getValue('PCN_DATEVALIDITE') <= DCP) AND (T1.getValue('PCN_DATEFIN') = DCP) then
      Begin
      If (LcTypeCp<>'') and (T1.getValue('PCN_TYPECONGE') <> LcTypeCp )then
        Begin
        T1 :=  Lc_tobCP.FindNext([''],[''],False);
        continue;
        End;
      if ChgePar then
        T1.ChangeParent(Result, -1)
      else
        Begin
        T2 := Tob.create('Les pris payés',Result,-1);
        T2.Dupliquer(T1,True,True,False);
        End;
      End;
    T1 :=  Lc_tobCP.FindNext([''],[''],False);
    End;
  If Result.detail.count = 0 then FreeAndNil(Result)
  else
     Result.detail.Sort('PCN_DATEVALIDITE');
End;



Function PGRecupTob_Acquis(Lc_tobCP : Tob; DCP : TDateTime; ChgePar : Boolean = False) : Tob;
Var
T1, T2 : Tob;
Begin
  Result := Tob.create('Les acquis',nil,-1);
  T1 :=  Lc_tobCP.FindFirst([''],[''],False);
  While Assigned(T1) do
    Begin
    If (T1.getValue('PCN_CODETAPE') <> 'C') AND (T1.getValue('PCN_CODETAPE') <> 'S')
    AND ( ( ((T1.getValue('PCN_TYPECONGE')='ACQ') OR (T1.getValue('PCN_TYPECONGE')='ACA')
    OR (T1.getValue('PCN_TYPECONGE')='ACS')) AND (T1.getValue('PCN_DATEFIN') <> DCP) )
     OR (T1.getValue('PCN_TYPECONGE')='ARR') OR (T1.getValue('PCN_TYPECONGE')='REP')  OR
     ( ((T1.getValue('PCN_TYPECONGE')='AJU') OR (T1.getValue('PCN_TYPECONGE')='AJP') or(T1.getValue('PCN_TYPECONGE')='REL'))     AND (T1.getValue('PCN_SENSABS') = '+')))
     AND (T1.getValue('PCN_DATEVALIDITE') <= DCP) then
      Begin
      if ChgePar then
        T1.ChangeParent(Result, -1)
      else
        Begin
        T2 := Tob.create('Les acquis',Result,-1);
        T2.Dupliquer(T1,True,True,False);
        End;
      End;
    T1 :=  Lc_tobCP.FindNext([''],[''],False);
    End;
  If Result.detail.count = 0 then FreeAndNil(Result)
  else
     Result.detail.Sort('PCN_DATEVALIDITE;PCN_TYPECONGE');
End;


Function PGRecupTob_AjuAjp(Lc_tobCP : Tob; DCP : TDateTime; ChgePar : Boolean = False) : Tob;
Var
T1, T2 : Tob;
Begin
  Result := Tob.create('Les ajustements',nil,-1);
  T1 :=  Lc_tobCP.FindFirst([''],[''],False);
  While Assigned(T1) do
    Begin
    IF  (T1.getValue('PCN_CODETAPE') <> 'C') AND (T1.getValue('PCN_CODETAPE') <> 'S')
    AND (( (T1.getValue('PCN_TYPECONGE')='AJP') OR (T1.getValue('PCN_TYPECONGE')='AJU'))
    AND (T1.getValue('PCN_SENSABS') = '-'))
    AND (T1.getValue('PCN_DATEVALIDITE') <= DCP) then
      Begin
      if ChgePar then
        T1.ChangeParent(Result, -1)
      else
        Begin
        T2 := Tob.create('Les ajustements',Result,-1);
        T2.Dupliquer(T1,True,True,False);
        End;
      End;
    T1 :=  Lc_tobCP.FindNext([''],[''],False);
    End;
  If Result.detail.count = 0 then FreeAndNil(Result)
  else
     Result.detail.Sort('PCN_DATEVALIDITE;PCN_TYPECONGE');
End;

procedure PgReintegreTobMere(LcTob,LcTob_GblCp : Tob);
Var T1 : Tob;
Begin
if not assigned(LcTob) then exit;
if not assigned(LcTob_GblCp) then exit;
T1 :=  LcTob.FindFirst([''],[''],False);
While Assigned(T1) do
  Begin
  T1.ChangeParent(LcTob_GblCp, -1);
  T1 :=  LcTob.FindNext([''],[''],False);
  End;
End;





Function PGExistSoldeCloture(Lc_tobCP: tob; DCP : TDateTime) : Boolean;
var
  TPris,T: tob;
Begin
  Result := False;
  if Lc_tobCP = nil then exit;
  TPris := PGRecupTob_prisPayes(Lc_tobCP,DCP,'SLD');
  if assigned(TPris) then
    T := TPris.findfirst(['PCN_TYPECONGE','PCN_GENERECLOTURE'], ['SLD','X'], True)
  else
    T := Nil;
  if Assigned(T) then Result := True;
End;


function PGCalculValoCP(LCTob_GblCp, LCTobEtab, LCTobSAL : tob; LcNbJOuvres, LcNbJOuvrables: double; LcDatedeb, Lcdatefin: tdatetime; var LcNext: integer; LcTypeCp: string): boolean;
var
  LcTb_pris,LcTb_delta,LcTb_Acquis,Tob_ElementCP, TP, TD : Tob;
  DiffRel, DiffPer1, DiffPer0, AcqEncours, ValoDxmn,HVal,DX : Double;
  DTcloture,DDPer1, DFPer1, DDPer0, Datedebut,FinMoisprec,Debmoisprec : TDateTime;
  LcJOuvres, LcJOuvrables, NbJoursReference, Nbjourspri, Cumul12, MS, IndemJourRel : double;
  ModeValo,MValoMs,PaieSalMaint, St : String;
  NbJSem, NoP : Integer;
  aa, mm, jj : Word;
  ValoJX : array[0..10] of double;
begin
  { Typecp = 'PRI' ou 'SLD' }
  result := true;
  If LCTob_GblCp = nil then exit;
  LcTb_pris := PGRecupTob_pris(LCTob_GblCp,Lcdatefin,'');
  if LcTb_pris = nil then exit;

  DTcloture := LCTobEtab.getvalue('ETB_DATECLOTURECPN');
  { ne doit jamais arriver }
  if DTcloture = 0 then exit;

  ModeValo := TrouveModeValorisation(LCTobEtab, LCTobSAL);

  PaieSalMaint := TrouvePaieSalMaintien(LCTobEtab, LCTobSAL);

  { basé sur calendrier réel ou théorique }
  RendMethodeValoDixieme(NBjSem, MValoMs, ValoDxmn);
  if MValoMs = 'T' then { Calendrier théorique }
  begin
    LcJOuvres := 21.67;
    LcJOuvrables := 26;
  end
  else
  begin
    if MValoMs = 'M' then { Gestion manuelle }
    begin
      LcJOuvres := ValoDxmn;
      LcJOuvrables := ValoDxmn;
    end
    else
    begin { Calendrier réel }
      if LcTypeCp = 'PRI' then
      begin
        LcJOuvres := LcNBJOuvres;
        LcJOuvrables := LcNBJOuvrables;
      end
      else
      begin
        CalculVarOuvresOuvrablesMois(LCTobEtab, nil, nil, DebutdeMois(LcDatedeb), FindeMois(LcDatedeb), TRUE, LcJOuvres, HVal);
        CalculVarOuvresOuvrablesMois(LCTobEtab, nil, nil, DebutdeMois(LcDatedeb), FindeMois(LcDatedeb), FALSE, LcJOuvrables, HVal);
      end;
    end;
  end;

  Datedebut := DebutDeMois(LcDateFin);
  FinMoisprec := datedebut - 1;
  Debmoisprec := DebutdeMois(FinMoisPrec);
  Decodedate(DTcloture, aa, mm, jj);

  DDPer0 := PGEncodeDateBissextile(aa - 1, mm, jj) + 1;
  DDPer1 := PGEncodeDateBissextile(aa - 2, mm, jj) + 1;
  DFPer1 := PGEncodeDateBissextile(aa - 1, mm, jj);

  { Recherche des éléments pour le calcul du PRI ou du SLD }
  Tob_ElementCP := PGCalculCompteurCp(LcTob_GblCp);
  if Tob_ElementCP = nil then exit;

  LcTb_pris   := PGRecupTob_pris(LCTob_GblCp,Lcdatefin,LcTypeCp,True);
  LcTb_Acquis := PGRecupTob_Acquis(LCTob_GblCp,Lcdatefin,True);
  LcTb_delta := PGRecupTob_AjuAjp(LCTob_GblCp,Lcdatefin,True);

  NbjSem := LCTobEtab.getvalue('ETB_NBJOUTRAV');
  if LcTobSal.GetValue('PSA_CPACQUISMOIS') = 'ETB' then
    NbJoursReference := JoursReference(LCTobEtab.getvalue('ETB_NBREACQUISCP'))
  else
    NbJoursReference := JoursReference(LcTobSal.getvalue('PSA_NBREACQUISCP'));


  for NoP := 0 to 1 do
    ValoJX[NoP] := CalculValoX(DTcloture, Lcdatefin, Lcdatedeb, NoP, NbJoursReference, LcTypecp, LcTb_Acquis, LcTb_delta);


  TP := LcTb_pris.findfirst(['PCN_TYPECONGE'], [LcTypeCp], True);
  while TP <> nil do
    begin
    if (TP.GetValue('PCN_MVTDUPLIQUE')='X') then
      Begin
      TP := LcTb_pris.findnext([''], [''], True);
      Continue;
      End;
    Nbjourspri := TP.getvalue('PCN_JOURS');
    DiffRel  := Tob_ElementCP.GetValue('RELIQUAT') - Tob_ElementCP.GetValue('PRIRELIQUAT');
    DiffPer1 := Tob_ElementCP.GetValue('ACQUISN1') - Tob_ElementCP.GetValue('PRIN1');
    DiffPer0 := Tob_ElementCP.GetValue('ACQUISN')  - Tob_ElementCP.GetValue('PRIN');
    //    BCPACQUIS, // Base acquis ds la session
    if (Nbjourspri > DiffRel + DiffPer1 + DiffPer0 ) then
      Begin
      result := False;
      TP := LcTb_pris.findnext(['PCN_TYPECONGE'], [LcTypeCp], True);
      Continue;
      End;
    result := True;
    TP.PutValue('PCN_VALOMS',0);
    TP.PutValue('PCN_VALOX',0) ;
    TP.PutValue('PCN_VALORETENUE',0);
    TP.PutValue('PCN_ABSENCE',0);
    TP.PutValue('PCN_BASE',0);
    if TP.GetValue('PCN_TYPECONGE') = 'PRI' then  St := 'Congés pris '
    else
      Begin
      St := 'Solde ';
      TP.PutValue('PCN_LIBELLE',St + FloatToStr(Nbjourspri) + 'j au '+DateToStr(LcDateFin));
      End;

    { PT108 Recherche du salaire perçu précedant la date de début d'absence et selon le paramétrage défini }
    Cumul12 := RecupSalMaintien (PaieSalMaint,TP.GetValue('PCN_DATEDEBUTABS'),LcDateDeb);

    { Traitement des reliquats }
    IF ((DiffRel) > 0) AND (Nbjourspri > 0) then
       Begin
       MS := 0;
       IndemJourRel := CalculValoReliquat(LcTb_Acquis, LcTb_delta, NbJoursReference);
       { Affectation des valeurs générales }
       TP.PutValue('PCN_PERIODEPY', 1);
       TP.PutValue('PCN_PERIODECP', 1);
       TP.putvalue('PCN_PERIODEPAIE', 'Reliquat au '+DateToStr(DFPer1));
       TP.PutValue('PCN_CODETAPE', 'P');
       TP.PutValue('PCN_DATEPAIEMENT', LcDateFin);
       if VH_Paie.PGEcabMonoBase then TP.PutValue('PCN_EXPORTOK','X');

       If Nbjourspri <= DiffRel then
         Begin  { Intégration du mvt en totalité sur le reliquat }
         DX := IndemJourRel * Nbjourspri;
         if (NbJSem = 5) and (LcJouvres <> 0) then MS := cumul12 * Nbjourspri / LcJouvres;
         if (NbJsem = 6) and (LcJouvrables <> 0) then MS := cumul12 * Nbjourspri / LcJouvrables;
         EcritMvtPris(TP, Ms, DX , Ms, ModeValo,LcDatedeb,LcDateFin);
         Tob_ElementCP.PutValue('PRIRELIQUAT',Tob_ElementCP.GetValue('PRIRELIQUAT') + Nbjourspri);
         Nbjourspri := 0;
         End
       else
         Begin  { Intégration du mvt partiellement sur le reliquat }
         { Duplication du mvt pour générer un mvt sur reliquat }
         TD := TOB.Create('ABSENCESALARIE', LcTb_pris, (TP.GetIndex) + 1);
         TD.Dupliquer(TP, FALSE, TRUE, TRUE);
         TD.putvalue('PCN_ORDRE', LcNext);
         LcNext := LcNext + 1;
         TD.Putvalue('PCN_CODETAPE', 'P');
         if VH_Paie.PGEcabMonoBase then TD.putvalue('PCN_EXPORTOK','X');
         TD.PutValue('PCN_MVTDUPLIQUE', 'X');
         if TP.getvalue('PCN_TYPECONGE') = 'PRI' then
         TD.PutValue('PCN_DATEPAIEMENT', LcDateFin);
         TD.PutValue('PCN_CODERGRPT', TP.getvalue('PCN_ORDRE'));
         TD.putvalue('PCN_JOURS', Arrondi(DiffRel, DCP));
         TD.putvalue('PCN_PERIODEPAIE', 'Reliquat au '+DateToStr(DFPer1));
         TD.PutValue('PCN_LIBELLE',St + FloatToStr(DiffRel) + 'j au '+DateToStr(LcDateFin));
         DX := IndemJourRel * DiffRel;
         if (NbJSem = 5) and (LcJouvres <> 0) then MS := cumul12 * DiffRel / LcJouvres;
         if (NbJsem = 6) and (LcJouvrables <> 0) then MS := cumul12 * DiffRel / LcJouvrables;
         EcritMvtPris(TD, Ms, DX, Ms, ModeValo,LcDatedeb,LcDateFin);
         EcritMvtPris(TP, (TP.GetValue('PCN_VALOMS') + Ms), (TP.GetValue('PCN_VALOX') + DX), Ms, ModeValo,LcDatedeb,LcDateFin);
         TP.PutValue('PCN_CODERGRPT', -1);
         Tob_ElementCP.PutValue('PRIRELIQUAT',Tob_ElementCP.GetValue('PRIRELIQUAT') + DiffRel);
         Nbjourspri := Nbjourspri - DiffRel;
         End;
       End;


     IF ((DiffPer1) > 0) AND (Nbjourspri > 0) then
       Begin
       MS := 0;
       TP.putvalue('PCN_PERIODEPAIE', 'Période du ' + datetostr(DDPer1) + ' au ' + datetostr(DFPer1));
       TP.putvalue('PCN_DATEPAIEMENT', LcDateFin);
       TP.putValue('PCN_PERIODEPY', 1);
       TP.PutValue('PCN_PERIODECP', 1);
       TP.PutValue('PCN_CODETAPE', 'P');
       if VH_Paie.PGEcabMonoBase then TP.putvalue('PCN_EXPORTOK','X');

       If Nbjourspri <= DiffPer1 then
          Begin  { Intégration du mvt en totalité sur la periode n-1 }
          DX := ValoJX[1] * Nbjourspri;
          if (NbJSem = 5) and (LcJouvres <> 0) then MS := cumul12 * Nbjourspri / LcJouvres;
          if (NbJsem = 6) and (LcJouvrables <> 0) then MS := cumul12 * Nbjourspri / LcJouvrables;
          EcritMvtPris(TP, (TP.GetValue('PCN_VALOMS') + Ms), (TP.GetValue('PCN_VALOX') + DX), Ms, ModeValo,LcDatedeb,LcDateFin);
          if (TP.GetValue('PCN_CODERGRPT')='-1') then
            Begin    { Si PRI Parent alors mvt dupliqué }
            TD := TOB.Create('ABSENCESALARIE', LcTb_pris, (TP.GetIndex) + 1);
            TD.Dupliquer(TP, FALSE, TRUE, TRUE);
            TD.putvalue('PCN_ORDRE', LcNext);
            LcNext := LcNext + 1;
            TD.Putvalue('PCN_CODETAPE', 'P');
            if VH_Paie.PGEcabMonoBase then TD.putvalue('PCN_EXPORTOK','X');
            TD.PutValue('PCN_MVTDUPLIQUE', 'X');
            TD.PutValue('PCN_DATEPAIEMENT', LcDateFin);
            TD.PutValue('PCN_CODERGRPT', TP.getvalue('PCN_ORDRE'));
            TD.putvalue('PCN_JOURS', Arrondi(Nbjourspri, DCP));
            EcritMvtPris(TD, Ms, DX, Ms, ModeValo,LcDatedeb,LcDateFin);
            TD.PutValue('PCN_LIBELLE',St + FloatToStr(Nbjourspri) + 'j au '+DateToStr(LcDateFin));
            TP.PutValue('PCN_CODERGRPT', -1);
            End;
          Tob_ElementCP.PutValue('PRIN1',Tob_ElementCP.GetValue('PRIN1') + Nbjourspri);
          Nbjourspri := 0;
          End
       else
          Begin  { Intégration du mvt partiellement sur la periode n-1 }
          TD := TOB.Create('ABSENCESALARIE', LcTb_pris, (TP.GetIndex) + 1);
          TD.Dupliquer(TP, FALSE, TRUE, TRUE);
          TD.putvalue('PCN_ORDRE', LcNext);
          LcNext := LcNext + 1;
          TD.Putvalue('PCN_CODETAPE', 'P');
          if VH_Paie.PGEcabMonoBase then TD.putvalue('PCN_EXPORTOK','X');
          TD.PutValue('PCN_MVTDUPLIQUE', 'X');
          if TP.getvalue('PCN_TYPECONGE') = 'PRI' then
          TD.PutValue('PCN_DATEPAIEMENT', LcDateFin);
          TD.PutValue('PCN_CODERGRPT', TP.getvalue('PCN_ORDRE'));
          TD.putvalue('PCN_JOURS', Arrondi(DiffPer1, DCP));
          TD.PutValue('PCN_LIBELLE',St + FloatToStr(Nbjourspri) + 'j au'+DateToStr(LcDateFin));
          DX := ValoJX[1] * DiffPer1;
          if (NbJSem = 5) and (LcJouvres <> 0) then MS := cumul12 * DiffPer1 / LcJouvres;
          if (NbJsem = 6) and (LcJouvrables <> 0) then MS := cumul12 * DiffPer1 / LcJouvrables;
          EcritMvtPris(TD, Ms, DX, Ms, ModeValo,LcDatedeb,LcDateFin);
          EcritMvtPris(TP, (TP.GetValue('PCN_VALOMS') + Ms), (TP.GetValue('PCN_VALOX') + DX), Ms, ModeValo,LcDatedeb,LcDateFin);
          TP.PutValue('PCN_CODERGRPT', -1);
          Tob_ElementCP.PutValue('PRIN1',Tob_ElementCP.GetValue('PRIN1') + DiffPer1);
          Nbjourspri := Nbjourspri - DiffPer1;
          End;
       End;


     IF ((DiffPer0 + AcqEncours ) > 0) AND (Nbjourspri > 0) then
       Begin
       MS := 0;
       TP.putvalue('PCN_PERIODEPAIE', 'Période du ' + datetostr(DDPer0) + ' au ' + datetostr(DTcloture));
       TP.putvalue('PCN_DATEPAIEMENT', Lcdatefin);
       TP.putValue('PCN_PERIODEPY', 0);
       TP.PutValue('PCN_PERIODECP', 0);
       TP.PutValue('PCN_CODETAPE', 'P');
       if VH_Paie.PGEcabMonoBase then TP.putvalue('PCN_EXPORTOK','X');

       If (Nbjourspri <= DiffPer0 + AcqEncours) then
          Begin  { Intégration du mvt en totalité sur la periode n }
          DX := ValoJX[0] * Nbjourspri;
          if (NbJSem = 5) and (LcJouvres <> 0) then MS := cumul12 * Nbjourspri / LcJouvres;
          if (NbJsem = 6) and (LcJouvrables <> 0) then MS := cumul12 * Nbjourspri / LcJouvrables;
          EcritMvtPris(TP, (TP.GetValue('PCN_VALOMS') + Ms), (TP.GetValue('PCN_VALOX') + DX), Ms, ModeValo,LcDatedeb,LcDateFin);
          if (TP.GetValue('PCN_CODERGRPT')='-1') then
            Begin    { Si PRI Parent alors mvt dupliqué }
            TD := TOB.Create('ABSENCESALARIE', LcTb_pris, (TP.GetIndex) + 1);
            TD.Dupliquer(TP, FALSE, TRUE, TRUE);
            TD.putvalue('PCN_ORDRE', LcNext);
            LcNext := LcNext + 1;
            TD.Putvalue('PCN_CODETAPE', 'P');
            if VH_Paie.PGEcabMonoBase then TD.putvalue('PCN_EXPORTOK','X');
            TD.PutValue('PCN_MVTDUPLIQUE', 'X');
            TD.PutValue('PCN_DATEPAIEMENT', LcDatefin);
            TD.PutValue('PCN_CODERGRPT', TP.getvalue('PCN_ORDRE'));
            TD.putvalue('PCN_JOURS', Arrondi(Nbjourspri, DCP));
            TD.PutValue('PCN_LIBELLE',St + FloatToStr(Nbjourspri) + 'j au '+DateToStr(LcDateFin));
            EcritMvtPris(TD, Ms, DX, Ms, ModeValo,LcDatedeb,LcDateFin);
            TP.PutValue('PCN_CODERGRPT', -1);
            if VH_Paie.PGEcabMonoBase then TP.putvalue('PCN_EXPORTOK','X');
            End;
          Tob_ElementCP.PutValue('PRIN',Tob_ElementCP.GetValue('PRIN') + Nbjourspri);
          //Nbjourspri := 0;
          End
       else
          Begin  { Intégration du mvt partiellement sur la periode n }
          TD := TOB.Create('ABSENCESALARIE', LcTb_pris, (TP.GetIndex) + 1);
          TD.Dupliquer(TP, FALSE, TRUE, TRUE);
          TD.putvalue('PCN_ORDRE', LcNext);
          LcNext := LcNext + 1;
          TD.Putvalue('PCN_CODETAPE', 'P');
          if VH_Paie.PGEcabMonoBase then TD.putvalue('PCN_EXPORTOK','X');
          TD.PutValue('PCN_MVTDUPLIQUE', 'X');
          if TP.getvalue('PCN_TYPECONGE') = 'PRI' then
          TD.PutValue('PCN_DATEPAIEMENT', LcDateFin);
          TD.PutValue('PCN_CODERGRPT', TP.getvalue('PCN_ORDRE'));
          TD.putvalue('PCN_JOURS', Arrondi(DiffPer0 + AcqEncours, DCP));
          TD.PutValue('PCN_LIBELLE',St + FloatToStr(Arrondi(DiffPer0 + AcqEncours, DCP)) + ' au '+DateToStr(LcDateFin));
          DX := ValoJX[0] * Arrondi(DiffPer0 + AcqEncours, DCP);
          if (NbJSem = 5) and (LcJouvres <> 0) then MS := cumul12 * Arrondi(DiffPer0 + AcqEncours, DCP) / LcJouvres;
          if (NbJsem = 6) and (LcJouvrables <> 0) then MS := cumul12 * Arrondi(DiffPer0 + AcqEncours, DCP) / LcJouvrables;
          EcritMvtPris(TD, Ms, DX, Ms, ModeValo,LcDatedeb,LcDateFin);
          EcritMvtPris(TP, (TP.GetValue('PCN_VALOMS') + Ms), (TP.GetValue('PCN_VALOX') + DX), Ms, ModeValo,LcDatedeb,LcDateFin);
          TP.PutValue('PCN_CODERGRPT', -1);
          if VH_Paie.PGEcabMonoBase then TP.putvalue('PCN_EXPORTOK','X');
          Tob_ElementCP.PutValue('PRIN',Tob_ElementCP.GetValue('PRIN') + DiffPer0 + AcqEncours);
          //Nbjourspri := Nbjourspri - (DiffPer0 + AcqEncours);
          End;
       End;

    TP := LcTb_pris.findnext(['PCN_TYPECONGE'], [LcTypeCp], True);
    End;


   AgregeCumulPris(LcTb_pris);

   PgReintegreTobMere(LcTb_pris,LCTob_GblCp);
   PgReintegreTobMere(LcTb_Acquis,LCTob_GblCp);
   PgReintegreTobMere(LcTb_delta,LCTob_GblCp);

end;


procedure PgGenereSolde(Tob_Sal, Tob_etab, LCTob_GblCp : tob; DateF, DateD: tdatetime; var next: integer; Action: string);
var
  DateSortie, DTFinP, DTDebP: TDateTime;
//  NbJAcq,  NBJPayes, NBJAcqP, cpt: double;
  T: Tob;
  Periode: integer;
  salarie: string;
  Tob_ElementCP: Tob;
  NbjAArr, NbjASld, Arr : double;
begin
  Salarie := Tob_Sal.GetValue('PSA_SALARIE');
  DateSortie := Tob_Sal.GetValue('PSA_DATESORTIE');
  if not ((DateSortie >= DateD) and (DateSortie <= Datef)) then exit;
  //NbJAcq := 0;
  //NBJPayes := 0;
  //cpt := 0;
  Periode := RendPeriode(DTFinP, DTDebP, Tob_etab.getvalue('ETB_DATECLOTURECPN'), DateF);

  Tob_ElementCP := PGCalculCompteurCp(LcTob_GblCp);

  If not Assigned(Tob_ElementCP) then exit;


  NbjAArr := Tob_ElementCP.GetValue('ACQUISN') - Tob_ElementCP.GetValue('PRIN');

  Arr := 0;  { PT111 SI CHR pas d'arrondi de solde }
  if VH_Paie.PGChr6Semaine = False then   { PT111 }
     Begin
     {DEB PT-11-2 //PT-24 On arrondi l'acquis de la periode N et non le total des 2 périodes (ASOLDERAVBULL) -JCPACQUISACONSOMME }
     Arr := Arrondi((1 + Int(NbjAArr) - (NbjAArr)), dcp);
     if ((Arr > 0) and (Arr < 1)) then
      begin
      T := TOB.Create('ABSENCESALARIE', LCTob_GblCp, -1);
      Next := Next + 1;
      CreeArrondi(Arr, Tob_Sal.getvalue('PSA_SALARIE'), tob_sal.getvalue('PSA_ETABLISSEMENT'), '-', Next, Periode, DateF, T); //mvi
      end
      else Arr := 0;
      { FIN PT-11-2 }
      End;



  NbjASld := Tob_ElementCP.GetValue('RELIQUAT') - Tob_ElementCP.GetValue('PRIRELIQUAT');
  NbjASld := NbjASld + Tob_ElementCP.GetValue('ACQUISN1') - Tob_ElementCP.GetValue('PRIN1');
  NbjASld := NbjASld + Tob_ElementCP.GetValue('ACQUISN')  - Tob_ElementCP.GetValue('PRIN');

  if ACtion = 'C' then
  begin
    if Arr = 0 then Next := Next + 1; { PT-21 Le num ordre doit être incrementé.. }
    T := TOB.Create('ABSENCESALARIE', LCTob_GblCp, -1);
    InitialiseTobAbsenceSalarie(T);
    T.PutValue('PCN_SALARIE', Tob_Sal.getvalue('PSA_SALARIE'));
    T.PutValue('PCN_ORDRE', Next);
    T.PutValue('PCN_TYPECONGE', 'SLD');
    T.PutValue('PCN_SENSABS', '-');
    T.PutValue('PCN_LIBELLE', 'Solde ' + floattostr(Arrondi(NbjASld + Arr, DCP)) + 'j au ' + datetostr(dateF)); //PT-11-1
    T.PutValue('PCN_DATEMODIF', Date);
    T.PutValue('PCN_DATEVALIDITE', Datef);
    T.PutValue('PCN_PERIODECP', 0);
    //T.PutValue('PCN_DATEDEBUT', Datef);
    //T.PutValue('PCN_DATEFIN', Datef);
    T.PutValue('PCN_JOURS', Arrondi(NbjASld + Arr, DCP)); //PT-11-1
    T.putValue('PCN_CODETAPE', '...'); { PT-13 }
    T.PutValue('PCN_ETABLISSEMENT', Tob_Sal.getvalue('PSA_ETABLISSEMENT'));
    T.Putvalue('PCN_TRAVAILN1', Tob_Sal.getvalue('PSA_TRAVAILN1'));
    T.Putvalue('PCN_TRAVAILN2', Tob_Sal.getvalue('PSA_TRAVAILN2'));
    T.Putvalue('PCN_TRAVAILN3', Tob_Sal.getvalue('PSA_TRAVAILN3'));
    T.Putvalue('PCN_TRAVAILN4', Tob_Sal.getvalue('PSA_TRAVAILN4'));
    T.Putvalue('PCN_CODESTAT', Tob_Sal.getvalue('PSA_CODESTAT'));
    T.Putvalue('PCN_CONFIDENTIEL', Tob_Sal.getvalue('PSA_CONFIDENTIEL'));
    Next := Next + 1;
    { PT-29 }
  end;
 (* if Action = 'M' then
  begin
    T := LCTob_GblCp.findfirst([''], [''], true);
    while t <> nil do
    begin
      if ((t.getvalue('PCN_CODERGRPT') = -1) or (LCTob_GblCp.detail.count = 1)) then
      begin
        T.PutValue('PCN_LIBELLE', 'Solde ' + floattostr(Arrondi(ASOLDERAVBULL + AcquisBulletin + arr, DCP)) + 'j au ' + datetostr(dateF));
        T.PutValue('PCN_JOURS', Arrondi(ASOLDERAVBULL + AcquisBulletin + arr, DCP)); //PT-11-2
        T := LCTob_GblCp.findnext([''], [''], true);
        continue;
      end;
      if ((t.getvalue('PCN_PERIODECP') <> 0)) then
      begin
        cpt := cpt + T.getValue('PCN_JOURS');
        T.PutValue('PCN_LIBELLE', 'Solde ' + floattostr(Arrondi(ASOLDERAVBULL + AcquisBulletin {mv  05-02+arr}, DCP)) + 'j au ' + datetostr(dateF));
      end;
      T := LCTob_GblCp.findnext([''], [''], true);
    end;
    T := LCTob_GblCp.findfirst(['PCN_PERIODEPY'], ['0'], true);
    while t <> nil do
    begin
      if t.getvalue('PCN_CODERGRPT') <> -1 then
      begin
        T.PutValue('PCN_LIBELLE', 'Solde ' + floattostr(Arrondi(ASOLDERAVBULL + AcquisBulletin + Arr, DCP)) + 'j au ' + datetostr(dateF));
        T.PutValue('PCN_JOURS', Arrondi(ASOLDERAVBULL + AcquisBulletin - cpt + arr, DCP));
      end;
      T := LCTob_GblCp.findnext(['PCN_PERIODEPY'], ['0'], true);
    end;
  end;    *)


End;


Procedure PgGetJHCPPris( LCTob_GblCp : Tob; DCP : TDateTime; Var Nbj, Nbh : double) ;
Var TPris,T1 : Tob;
Begin
  Nbj := 0;
  Nbh := 0;
  if not assigned(LCTob_GblCp) then exit;
  TPris := PGRecupTob_prisPayes(LCTob_GblCp,DCP,'PRI');
  if not assigned(TPris) then exit;
  NbJ := TPris.Somme('PCN_JOURS',['PCN_MVTDUPLIQUE'],['-'],False);
  Nbh := TPris.Somme('PCN_HEURES',['PCN_MVTDUPLIQUE'],['-'],False);
End;


Procedure  PgGetJHCPPrisPoses( LCTob_GblCp : Tob; DCP : TDateTime; Var Nbj, Nbh : double) ;
Var TPris,T1 : Tob;
Begin
  Nbj := 0;
  Nbh := 0;
  if not assigned(LCTob_GblCp) then exit;
  TPris := PGRecupTob_pris(LCTob_GblCp,DCP,'PRI');
  if assigned(TPris) then
    Begin
    Nbj := TPris.Somme('PCN_JOURS',['PCN_MVTDUPLIQUE'],['-'],False);
    Nbh := TPris.Somme('PCN_HEURES',['PCN_MVTDUPLIQUE'],['-'],False);
    End;
  TPris := PGRecupTob_prisPayes(LCTob_GblCp,DCP,'PRI');
  if assigned(TPris) then
    Begin
    Nbj := Nbj + TPris.Somme('PCN_JOURS',['PCN_MVTDUPLIQUE'],['-'],False);
    Nbh := Nbh + TPris.Somme('PCN_HEURES',['PCN_MVTDUPLIQUE'],['-'],False);
    End;
End;


Procedure PgGetMTCPPris( LCTob_GblCp : Tob; DCP : TDateTime; Var Abs, Ind : double) ;
Var TPris,T1 : Tob;
Begin
  Abs := 0;
  Ind := 0;
  if not assigned(LCTob_GblCp) then exit;
  TPris := PGRecupTob_prisPayes(LCTob_GblCp,DCP,'PRI');
  if not assigned(TPris) then exit;
  T1 := TPris.FindFirst([''],[''],False);
  While Assigned(T1) do
    Begin
    if T1.getvalue('PCN_MVTDUPLIQUE') <> 'X' then
      Begin
      if T1.GetValue('PCN_MODIFABSENCE') = 'X' then
        Abs := Abs + T1.GetValue('PCN_ABSENCEMANU')
      else
        Abs := Abs + T1.GetValue('PCN_ABSENCE');
      if T1.Getvalue('PCN_MODIFVALO') <> 'X' then
        Ind := Ind + T1.GetValue('PCN_VALORETENUE')
      else
        Ind := Ind + T1.GetValue('PCN_VALOMANUELLE');
      End;
    T1 := TPris.FindNext([''],[''],False);
   End;
End;




Procedure  PgGetJHCPSolde( LCTob_GblCp : Tob; DCP : TDateTime; Var Nbj, Nbh : double) ;
Var TPris,T1 : Tob;
Begin
  Nbj := 0;
  Nbh := 0;
  if not assigned(LCTob_GblCp) then exit;
  TPris := PGRecupTob_prisPayes(LCTob_GblCp,DCP,'SLD');
  if not assigned(TPris) then exit;
  NbJ := TPris.Somme('PCN_JOURS',['PCN_MVTDUPLIQUE'],['-'],False);
  Nbh := TPris.Somme('PCN_HEURES',['PCN_MVTDUPLIQUE'],['-'],False);
End;


Procedure  PgGetMTCPSolde( LCTob_GblCp : Tob; DCP : TDateTime; Var Abs, Ind : double) ;
Var TPris,T1 : Tob;
Begin
  Abs := 0;
  Ind := 0;
  if not assigned(LCTob_GblCp) then exit;
  TPris := PGRecupTob_prisPayes(LCTob_GblCp,DCP,'SLD');
  if not assigned(TPris) then exit;
  T1 := TPris.FindFirst([''],[''],False);
  While Assigned(T1) do
    Begin
    if T1.getvalue('PCN_MVTDUPLIQUE') <> 'X' then
      Begin
      if T1.GetValue('PCN_MODIFABSENCE') = 'X' then
        Abs := Abs + T1.GetValue('PCN_ABSENCEMANU')
      else
        Abs := Abs + T1.GetValue('PCN_ABSENCE');
      if T1.Getvalue('PCN_MODIFVALO') <> 'X' then
        Ind := Ind + T1.GetValue('PCN_VALORETENUE')
      else
        Ind := Ind + T1.GetValue('PCN_VALOMANUELLE');
      End;
    T1 := TPris.FindNext([''],[''],False);
   End;
End;




Function   PgGetJCPArr ( LCTob_GblCp : Tob; DCP : TDateTime) : double ;
Begin
  Result := 0;
  if not assigned(LCTob_GblCp) then exit;
  Result :=  LCTob_GblCp.Somme('PCN_JOURS',['PCN_TYPECONGE','PCN_DATEVALIDITE','PCN_GENERECLOTURE'],
                                           ['ARR',DCP,'-'],False);
End;




 (*
{ Les Pris }
  st := 'SELECT * FROM ABSENCESALARIE WHERE' +
    ' PCN_SALARIE = "' + TobSal.getvalue('PSA_SALARIE') + '"' +
    ' AND PCN_TYPEMVT="CPA" ' +
  ' AND ((PCN_TYPECONGE = "PRI" AND PCN_VALIDRESP = "VAL") OR PCN_TYPECONGE = "SLD")' +
    ' AND PCN_DATEVALIDITE <= "' + usdatetime(DTFinP) + '"' +
    ' AND PCN_DATEFIN <= "' + usdatetime(10) + '"' +
    ' ORDER BY PCN_DATEVALIDITE';
{ Les acquis }
st := 'SELECT * from ABSENCESALARIE ' +
    'WHERE PCN_SALARIE = "' + Tob_SAL.GetValeur(iPSA_SALARIE) + '" ' +
    'AND PCN_TYPEMVT="CPA" ' +
    'AND PCN_CODETAPE <> "C" AND PCN_CODETAPE <> "S" AND ' +
  '( ( (PCN_TYPECONGE="ACQ" OR PCN_TYPECONGE="ACA" OR PCN_TYPECONGE="ACS")  ' +
    ' AND PCN_DATEFIN <> "' + usdatetime(Datef) + '")' +
    ' OR PCN_TYPECONGE="ARR" OR PCN_TYPECONGE="REP" OR PCN_TYPECONGE="REB" OR ' +
    ' ( ((PCN_TYPECONGE="AJU") OR (PCN_TYPECONGE="AJP") or(PCN_TYPECONGE="REL")) AND PCN_SENSABS = "+"))' +
    ' AND PCN_DATEVALIDITE <= "' + usdatetime(Datef) + '" ' +
    ' ORDER BY PCN_DATEVALIDITE, PCN_TYPECONGE';



{ Les ajusts }

  st := 'SELECT * FROM ABSENCESALARIE ' +
    'WHERE PCN_SALARIE = "' + SALARIE + '" AND PCN_TYPEMVT="CPA" ' +
  'AND PCN_CODETAPE <> "C" AND PCN_CODETAPE <> "S" ' +
  'AND (( PCN_TYPECONGE="AJP" OR PCN_TYPECONGE="AJU") AND PCN_SENSABS = "-") ' +
  'AND PCN_DATEVALIDITE <= "' + usdatetime(DateF) + '" '+
  'ORDER BY PCN_DATEVALIDITE, PCN_TYPECONGE';

  *)











end.
