{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 13/05/2004
Modifié le ... : 25 /07  /2005
Description .. : Source TOF de la FICHE : AMUNITEOEUVRE ()
Mots clefs ... : TOF;AMUNITEOEUVRE
Suite ........ : BTY 07/05 FQ 16285 Gérer montants selon nb décimales de la monnaie du dossier
*****************************************************************}
Unit AMUNITEOEUVRE_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Windows,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     eMul,
     MaineAGL,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     AGLInit,
     UTOF,
     uTOB,
     ImEnt ;

Type
  TOF_AMUNITEOEUVRE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
      fPlanUO : TOB;
      fGrille : THGrid;
      fTotalUO : double;
      fResteAAffecter : double;
      fbModifie : boolean;
      fCurMontant : double;
      procedure InitLaGrille;
      procedure CalculCumulUO ( T : TOB ) ;
      procedure RecupereGrille ;
      procedure OnGrilleCellExit(Sender: TObject; var ACol,
        ARow: Integer; var Cancel: Boolean);
      procedure OnGrilleCellEnter(Sender: TObject; var ACol,
        ARow: Integer; var Cancel: Boolean);
      procedure OnGrilleRowEnter(Sender: TObject;
        Ou: Integer; var Cancel: Boolean; Chg: Boolean);
      procedure OnGrilleKeyDown (Sender: TObject; var Key: Word; Shift: TShiftState);
      procedure RafraichirGrille(bLectureGrille: boolean);
      function SoldeUO ( Ligne : integer ) : double;
      procedure RepartirLigneSuivantes ( Ligne : integer);
  end ;

const
      UO_COL_DATE     = 0;
      UO_COL_NOMBRE   = 1;
      UO_COL_CUMUL    = 2;
      UO_COL_POURCENT = 3;

procedure AMLanceFiche_UniteOeuvre;

Implementation

procedure AMLanceFiche_UniteOeuvre;
begin
  AGLLanceFiche('AM','AMUNITEOEUVRE','','','');
end;

procedure TOF_AMUNITEOEUVRE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AMUNITEOEUVRE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AMUNITEOEUVRE.OnUpdate ;
// var T : TOB;
var VerifOK: Boolean;
    Sender: Tobject;
    NumColEnCours,NumRowEnCours : integer;
begin
  Inherited ;
  if (fResteAAffecter <> 0) then
  begin
    LastError := -1;
    LastErrorMsg := 'Des unités ne sont pas affectées correctement. Veuillez revoir la répartition.';
  end else
  begin
    // BTY 07/05 Modif cellule + souris sur bValider sans tabuler
    // provoque une répartition erronée sur les autres cellules
    // => prendre en compte cellule en cours avant de sortir
    Sender:=nil ; //YCP 21/10/05
    NumColEnCours := FGrille.Col;
    NumRowEnCours := FGrille.Row;
    VerifOK := False;
    OnGrilleCellExit (Sender,NumColEnCours,NumRowEnCours,VerifOK);
    if NOT VerifOK then
    begin
    //
      RecupereGrille;
      TheTOB := TOB.Create ('',nil,-1);
      TheTob.Dupliquer(fPlanUO,True,True);
//      TheTob.AddChampSupValeur('MODIFIE',fbModifie);
      Ecran.Close;
    end;
  end;
end ;

procedure TOF_AMUNITEOEUVRE.OnLoad ;
begin
  Inherited ;
  // BTY 07/05 Fait dans le OnArgument
  // InitLaGrille;
end ;

procedure TOF_AMUNITEOEUVRE.OnArgument (S : String ) ;
var i, Ou : integer;
begin
  Inherited ;
  { Evénements }
  fGrille := THGrid ( GetControl ('FListe') );
  fGrille.OnCellExit  := OnGrilleCellExit;
  fGrille.OnCellEnter  := OnGrilleCellEnter;
  fGrille.OnRowEnter := OnGrilleRowEnter;
  fGrille.OnKeyDown := OnGrilleKeyDown;
  // BTY 07/05 FQ 16285
  InitLaGrille;
  //
  fPlanUO := TOB.Create ('', nil, - 1);
  if LaTOB <> nil then fPlanUO.Dupliquer(LaTob,True,True);
  LaTOB := nil;

  fTotalUO := fPlanUO.GetValue('TOTALUO');

  fResteAAffecter := 0;

  RafraichirGrille ( False );

  fbModifie := False;

  { Positionnement sur la première ligne modifiable }
  Ou := -1;
  for i := 0 to fPlanUO.Detail.Count - 1 do
  begin
    if ((fPlanUO.Detail[i].GetValue('IUO_DATE') < VHImmo^.Encours.Deb) and (fPlanUO.GetValue('ETAT')<>'OUV')) then continue
    else
    begin
      Ou := i;
      break;
    end;
  end;
  { Dans un premier temps, dès qu'une immo a été clôturée, on ne modifie pas le plan d'unités d'oeuvre }
  if fPlanUO.GetValue('ETAT')<>'OUV' then Ou := -1;

  if Ou = -1 then fGrille.Enabled := False
  else fGrille.Row := Ou+1;

  fCurMontant := Valeur (fGrille.Cells[FGrille.Col,fGrille.Row]);

  SetControlVisible ('BVALIDER',(fPlanUO.GetValue('ETAT')='OUV'));
end ;

procedure TOF_AMUNITEOEUVRE.RafraichirGrille ( bLectureGrille : boolean );
var i : integer;
    stResteAAffecter : string;
begin
  fGrille.RowCount := fPlanUO.Detail.Count+1;
  if bLectureGrille then
  begin
    for i:=1 to fGrille.RowCount-1 do
      fPlanUO.Detail[i-1].PutValue('IUO_UNITEOEUVRE',Valeur(fGrille.Cells[UO_COL_NOMBRE,i]));
  end;
  CalculCumulUO ( fPlanUO );
  fPlanUO.PutGridDetail(THGrid(GetControl('FLISTE')),False,False,'IUO_DATE;IUO_UNITEOEUVRE;CUMULUO;POURCENTUO');
{*  // BTY Fiche 16285 Remplir fGrille avec les TOBs fPlanUO.Detail à formater sur nb décimales de la devise
for i:=1 to fGrille.RowCount-1 do
     begin
     fGrille.Cells[UO_COL_DATE,i] := fPlanUO.Detail[i-1].GetValue ('IUO_DATE');
     fGrille.Cells[UO_COL_NOMBRE,i] :=
     StrFMontant((fPlanUO.Detail[i-1].GetValue ('IUO_UNITEOEUVRE')),15 , V_PGI.OkDecV, '' , true);
     fGrille.Cells[UO_COL_CUMUL,i] :=
     StrFMontant((fPlanUO.Detail[i-1].GetValue ('CUMULUO')),15 , V_PGI.OkDecV, '' , true);
     fGrille.Cells[UO_COL_POURCENT,i] :=
     StrFMontant((fPlanUO.Detail[i-1].GetValue ('POURCENTUO')),15 , 2, '' , true);
     end;
  //
*}

  fResteAAffecter := Arrondi(fTotalUO - fPlanUO.Detail[fPlanUO.Detail.Count-1].GetValue('CUMULUO'),V_PGI.OkDecV);
  if fResteAAffecter > 0 then stResteAAffecter := StrFMontant(fResteAAffecter,15 , V_PGI.OkDecV, '' , true)+' '+TraduireMemoire('unités à répartir.')
  else if fResteAAffecter < 0 then stResteAAffecter := StrFMontant((-1)*fResteAAffecter,15 , V_PGI.OkDecV, '' , true)+TraduireMemoire(' unités à supprimer.')
  else stResteAAffecter := TraduireMemoire ('Repartition correcte.');
  SetControlCaption('RESTEAAFFECTER',stResteAAffecter);
end;

procedure TOF_AMUNITEOEUVRE.OnClose ;
begin
  fPlanUO.Free;
  Inherited ;
end ;

procedure TOF_AMUNITEOEUVRE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_AMUNITEOEUVRE.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_AMUNITEOEUVRE.CalculCumulUO(T: TOB);
var i : integer;
    CumulUO : double;
    TotalUO : double;
    Pourcent,SommePourcent : double;
begin
  CumulUO := 0;
  Pourcent := 0;
  SommePourcent := 0;
  TotalUO := Arrondi(T.Somme('IUO_UNITEOEUVRE',[''],[''],False), V_PGI.OkDecV);
  for i := 0 to T.Detail.Count - 1 do
  begin
    CumulUO := CumulUO + T.Detail[i].GetValue('IUO_UNITEOEUVRE');
    T.Detail[i].AddChampSupValeur('CUMULUO',CumulUO);
    if (TotalUO<>0) then Pourcent := Arrondi((T.Detail[i].GetValue('IUO_UNITEOEUVRE')/ TotalUO)*100, V_PGI.OkDecV);
    SommePourcent := Arrondi (SommePourcent + Pourcent, V_PGI.OkDecV);
    T.Detail[i].AddChampSupValeur('POURCENTUO',Pourcent );
  end;
  { Pour un total à 100 % }
  if T.Detail.Count > 0 then
  begin
    T.Detail[T.Detail.Count-1].PutValue('POURCENTUO',Arrondi(100-(SommePourcent-Pourcent), V_PGI.OkDecV));
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 19/05/2004
Modifié le ... :   /  /
Description .. : Initialisation de la grille
Suite ........ :  - typage des colonnes
Suite ........ :  - dimensionnement des colonnes
Mots clefs ... :
*****************************************************************}
procedure TOF_AMUNITEOEUVRE.InitLaGrille;
var LaGrille : THGrid;
    Masque : string;
begin
  LaGrille := THGrid ( GetControl ('FLISTE') );
  // BTY 07/05 FQ 16285 Gérer montants selon nb décimales de la monnaie du dossier
  Masque := StrfMask(V_PGI.OkDecV, '', True);


  LaGrille.ColTypes[UO_COL_DATE]:='D';
  LaGrille.ColFormats[UO_COL_DATE]:=ShortDateFormat;
  LaGrille.ColEditables[UO_COL_DATE]:=False;
  LaGrille.ColWidths[UO_COL_DATE]:=100;
  LaGrille.ColAligns[UO_COL_DATE]:=taCenter;
  LaGrille.Cells[UO_COL_DATE,0] := TraduireMemoire('Date');

  LaGrille.ColTypes[UO_COL_NOMBRE]:='R';
  //BTY 07/05 FQ 16285 LaGrille.ColFormats[UO_COL_NOMBRE]:='#,##0.00';
  LaGrille.ColFormats[UO_COL_NOMBRE]:=Masque;
  LaGrille.ColEditables[UO_COL_NOMBRE]:=True;
  LaGrille.ColWidths[UO_COL_NOMBRE]:=100;
  LaGrille.ColAligns[UO_COL_NOMBRE]:=taRightJustify;
  LaGrille.Cells[UO_COL_NOMBRE,0] := TraduireMemoire('Nombre');

  LaGrille.ColTypes[UO_COL_CUMUL]:='R';
  //BTY 07/05 FQ 16285 LaGrille.ColFormats[UO_COL_CUMUL]:='#,##0.00';
  LaGrille.ColFormats[UO_COL_CUMUL]:=Masque;
  LaGrille.ColEditables[UO_COL_CUMUL]:=False;
  LaGrille.ColWidths[UO_COL_CUMUL]:=100;
  LaGrille.ColLengths[UO_COL_CUMUL]:=-1;
  LaGrille.ColAligns[UO_COL_CUMUL]:=taRightJustify;
  LaGrille.Cells[UO_COL_CUMUL,0] := TraduireMemoire('Cumul');

  LaGrille.ColTypes[UO_COL_POURCENT]:='R';
  LaGrille.ColFormats[UO_COL_POURCENT]:='#,##0.00';
  LaGrille.ColEditables[UO_COL_POURCENT]:=False;
  LaGrille.ColWidths[UO_COL_POURCENT]:=100;
  LaGrille.ColLengths[UO_COL_POURCENT]:=-1;
  LaGrille.ColAligns[UO_COL_POURCENT]:=taRightJustify;
  LaGrille.Cells[UO_COL_POURCENT,0] := TraduireMemoire('%');
end;

procedure TOF_AMUNITEOEUVRE.OnGrilleCellExit(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
var RdMontant : double;
begin
  case ACol of
    UO_COL_NOMBRE :
      if IsNumeric(fGrille.Cells[ACol,ARow]) then
      begin
        RdMontant := Valeur(FGrille.Cells[ACol,ARow]);
        Cancel    := RdMontant < 0; // on ne peut pas saisir de valeur negative
        if Cancel then
          PGIInfo('Montant négatif !',ECRAN.Caption)
        else
        begin
          if (fCurMontant <> RdMontant) then
          begin
            FGrille.Cells[ACol,ARow] := STRFMONTANT ( RdMontant , 15 , V_PGI.OkDecV, '' , true);
            RepartirLigneSuivantes (ARow);
            RafraichirGrille ( True );
            fbModifie := True;
          end;
        end;
        fCurMontant := Valeur(FGrille.Cells[ACol,ARow]);
      end;
  end;
end;

procedure TOF_AMUNITEOEUVRE.RecupereGrille;
var i : integer;
begin
  for i := 0 to fPlanUO.Detail.Count - 1 do
  begin
    if IsNumeric(fGrille.Cells[UO_COL_NOMBRE,i+1]) then
    begin
      fPlanUO.Detail[i].PutValue('IUO_UNITEOEUVRE', Valeur(FGrille.Cells[UO_COL_NOMBRE,i+1]));
    end;
  end;
end;

procedure TOF_AMUNITEOEUVRE.OnGrilleKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  case Key of
    VK_F6 : fGrille.Cells[UO_COL_NOMBRE,fGrille.Row] := STRFMONTANT ( SoldeUO ( fGrille.Row ), 15 , V_PGI.OkDecV, '' , true);
  end;
end;

function TOF_AMUNITEOEUVRE.SoldeUO(Ligne: integer): double;
var i : integer;
    SommeUO : double;
begin
  SommeUO := 0;
  for i:=0 to fPlanUO.Detail.Count - 1 do
    SommeUO := SommeUO + fPlanUO.Detail[i].GetValue('IUO_UNITEOEUVRE');
  Result := fTotalUO - (SommeUO - fPlanUO.Detail[Ligne-1].GetValue('IUO_UNITEOEUVRE'));
  if Result < 0 then Result := 0;
end;

procedure TOF_AMUNITEOEUVRE.RepartirLigneSuivantes(Ligne: integer);
var
  nLigneARepartir : integer;
  i : integer;
  Montant, ResteARepartir : double;
begin
  nLigneARepartir := fGrille.RowCount - Ligne - 1;
  ResteARepartir := fTotalUO;
  for i := 1 to Ligne do
    ResteARepartir := ResteARepartir - Valeur(fGrille.Cells[UO_COL_NOMBRE,i]);
  if nLigneARepartir = 0 then exit;    
  Montant := Arrondi((ResteARepartir/nLigneARepartir),V_PGI.OkDecV);
  if Montant < 0 then exit;
  for i := Ligne+1 to fGrille.RowCount - 2 do
  begin
    fGrille.Cells[UO_COL_NOMBRE,i] := STRFMONTANT ( Montant , 15 , V_PGI.OkDecV, '' , true);
    ResteARepartir := ResteARepartir-Montant;
  end;
  if ResteARepartir < 0 then ResteARepartir := 0;
  fGrille.Cells[UO_COL_NOMBRE,fGrille.RowCount-1] := STRFMONTANT ( Arrondi(ResteARepartir,V_PGI.OkDecV) , 15 , V_PGI.OkDecV, '' , true);
end;

procedure TOF_AMUNITEOEUVRE.OnGrilleCellEnter(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
  fCurMontant := Valeur (fGrille.Cells[FGrille.Col,FGrille.Row]);
end;

procedure TOF_AMUNITEOEUVRE.OnGrilleRowEnter(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
  if ((fPlanUO.GetValue('ETAT')<>'OUV') and
      (fPlanUO.Detail[Ou-1].GetValue('IUO_DATE') < VHImmo^.Encours.Deb)) then
  Cancel := True
end;

Initialization
  registerclasses ( [ TOF_AMUNITEOEUVRE ] ) ;
end.
