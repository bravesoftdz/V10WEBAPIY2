unit planeche;

interface

uses
  (*Windows, (*Messages, *)SysUtils, Classes, (*Graphics, *)Controls, Forms, (*Dialogs,*)
  Grids, Hctrls, HPanel, StdCtrls, (*Buttons, (*ExtCtrls,*)HEnt1, (*ImEnt, *)hmsgbox,
  HTB97, (*UiUtil, *)HSysMenu, Menus, ExtCtrls;

function  FicheSaisieEcheancier (DateTranche : TDateTime;Periode : integer; Versement:string; var LTranche : TList) : boolean;
//function CalculNombreEcheance (DateDebut, DateFin : TDateTime; Periode : integer) : integer;  //YCP 25/08/05 déplacé

const E_DEBUT=1;
      E_FIN=2;
      E_NB=3;
      E_MONTANT=4;
      E_FRAIS=5;
type
  PTrancheEcheance = ^ATrancheEcheance;
  ATrancheEcheance = record
    DateDebut, DateFin: TDateTime;
    nEcheance : integer;
    Montant,Frais : double;
  end;
  PEcheance = ^AEcheance;
  AEcheance = record
    Date: TDateTime;
    Montant,Frais:double;
  end;
  TFPlanEche = class(TForm)
    HM1: THMsgBox;
    HPanel3: THPanel;
    HLabel3: THLabel;
    TotalValeur: THNumEdit;
    TotalFrais: THNumEdit;
    TotalNombre: THNumEdit;
    HMTrad: THSystemMenu;
    GridSaisie: THGrid;
    Dock971: TDock97;
    PBouton: TToolWindow97;
    BAbandon: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    BImprimer: TToolbarButton97;
    BValide: TToolbarButton97;
    PopupMenu: TPopupMenu;
    Inserer: TMenuItem;
    Supprimer: TMenuItem;
    HPanel1: THPanel;
    HLabel1: THLabel;
    HLabel2: THLabel;
    PERIODICITE: TEdit;
    VERSEMENT: TEdit;
    BInserer: TToolbarButton97;
    BSupprimer: TToolbarButton97;
    procedure GridSaisieCellExit(Sender: TObject; var ACol, ARow: Integer;
      var Cancel: Boolean);
    procedure FormShow(Sender: TObject);
    procedure GridSaisieRowExit(Sender: TObject; Ou: Integer;
      var Cancel: Boolean; Chg: Boolean);
    procedure GridSaisieCellEnter(Sender: TObject; var ACol, ARow: Integer;
      var Cancel: Boolean);
    procedure GridSaisieColExit(Sender: TObject; Ou: Integer;
      var Cancel: Boolean; Chg: Boolean);
    procedure GridSaisieRowEnter(Sender: TObject; Ou: Integer;
      var Cancel: Boolean; Chg: Boolean);
    procedure GridSaisieGetEditMask(Sender: TObject; ACol, ARow: Integer;
      var Value: String);
    procedure BValideClick(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure InsererClick(Sender: TObject);
    procedure SupprimerClick(Sender: TObject);
    procedure HelpBtnClick(Sender: TObject);
  private
    { Déclarations privées }
    Periode : integer;
    stCur : string;
    fVersement : string;
    procedure EnregistreListeTrancheEcheance ( var LTranche : TList);
    procedure InitGridSaisie(Date : TDateTime;LTranche : TList);
    procedure CalculCumul;
    function IsLigneVide (Row : integer): boolean;
    function IsLigneOK (Row : integer): boolean;
    function ControleCoherenceDate : boolean;    
  public
    { Déclarations publiques }
  end;


implementation

uses Outils
{$IFDEF eAGLCLient}
   ;
{$ELSE}
   , PrintDBG;
{$ENDIF}

{$R *.DFM}

function CalculDateFin(DateDebut: TDateTime; Periode,NbPeriode: integer) : TDateTime;
//var nbEcheance : integer; laDate : TDateTime;
begin
{  result:=idate1900 ;
  if NbPeriode<=0 then exit ;
  nbEcheance:=0;
  Dec(NbPeriode) ;
  laDate:=DateDebut;
  while nbEcheance<>NbPeriode do
    begin
    Inc(nbEcheance,1);
    laDate := PlusMois (laDate,Periode);
    laDate := FINDEMOIS(LaDate);
  end;
  result:=FINDEMOIS(LaDate);}
  Result := FINDEMOIS ( PlusMois ( DateDebut, (NbPeriode*Periode)-1 ) );
end;




function  FicheSaisieEcheancier (DateTranche : TDateTime;Periode : integer; Versement:string; var LTranche : TList) : boolean;
var
  FicheEcheancier: TFPlanEche;
begin
FicheEcheancier := TFPlanEche.Create(Application);
FicheEcheancier.Periode := Periode;
FicheEcheancier.fVersement := Versement;
try
FicheEcheancier.InitGridSaisie(DateTranche,LTranche);
if FicheEcheancier.ShowModal =  mrOk then
begin
  Result := True;
  FicheEcheancier.EnregistreListeTrancheEcheance (LTranche);
end else Result := False;
finally
FicheEcheancier.GridSaisie.VidePile (True);
FicheEcheancier.Free;
end;
end;

procedure TFPlanEche.GridSaisieCellEnter(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
  stCur := GridSaisie.Cells [ACol,Arow];
  case ACol of
    E_NB : begin
             if (GridSaisie.Cells[E_DEBUT,ARow]<>'') and (GridSaisie.Cells[E_FIN,ARow]<>'') then
                   GridSaisie.Cells[E_NB,ARow]:= IntToStr (CalculNombreEcheance (StrToDate(GridSaisie.Cells[E_DEBUT,ARow]),
                                                                                StrToDate(GridSaisie.Cells[E_FIN,ARow]), Periode));
             if GridSaisie.Cells[E_MONTANT,ARow] <> '' then exit;
             if (IsFloat(GridSaisie.Cells[E_MONTANT,ARow-1])) or
                (IsNumeric(GridSaisie.Cells[E_MONTANT,ARow-1])) then
                         GridSaisie.Cells[E_MONTANT,ARow]:= GridSaisie.Cells[E_MONTANT,ARow-1]
             else GridSaisie.Cells[E_MONTANT,ARow]:= MontantToStr(0.00);
             if GridSaisie.Cells[E_FRAIS,ARow] <> '' then exit;
             if (IsFloat(GridSaisie.Cells[E_FRAIS,ARow-1])) or
                (IsNumeric(GridSaisie.Cells[E_FRAIS,ARow-1])) then
                  GridSaisie.Cells[E_FRAIS,ARow]:= GridSaisie.Cells[E_FRAIS,ARow-1]
             else GridSaisie.Cells[E_FRAIS,ARow]:= MontantToStr(0.00);
            end;
    end;
end;

procedure TFPlanEche.GridSaisieCellExit(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
var Cellule : string;
begin
  Cellule := GridSaisie.Cells[ACol,ARow];
   if Cellule <> '' then
   begin
     if (Cellule <> stCur) or (ACol = 3) then
     case ACol of
       1,2 : begin
               if IsValidDate(Cellule) = false then Cancel := true
               else
               begin
                 GridSaisie.Cells[ACol,ARow] := DateToStr(StrToDate(Cellule));
                 if (IsValidDate (GridSaisie.Cells[1,ARow])) and (IsValidDate (GridSaisie.Cells[2,ARow]))
                   and ((StrToDate (GridSaisie.Cells[1,ARow])) >= (StrToDate(GridSaisie.Cells[2,ARow]))) then
                     HM1.Execute(0, Caption, '')
                 else if (IsValidDate (GridSaisie.Cells[1,ARow])) and (IsValidDate (GridSaisie.Cells[2,ARow])) then
                      GridSaisie.Cells[3,ARow] := IntToStr (CalculNombreEcheance (StrToDate(GridSaisie.Cells[1,ARow]),
                                                           StrToDate(GridSaisie.Cells[2,ARow]), Periode));
               end;
             end;
       3:  if IsNumeric (Cellule) = false then
             Cancel := true
           else if (GridSaisie.Cells[2,ARow] = '') or (stCur <> Cellule) then
             GridSaisie.Cells[2,ARow]:=DateToStr(CalculDateFin(StrToDate(GridSaisie.Cells[1,ARow]),Periode,StrToInt(GridSaisie.Cells[3,ARow]))) ;
               //EPZ 10/11/00              GridSaisie.Cells[2,ARow] :=DateToStr (PlusMois ( StrToDate(GridSaisie.Cells[1,ARow]), Periode* StrToInt(GridSaisie.Cells[3,ARow]))-1);
              //:=DateToStr (FINDEMOIS(PlusMois ( StrToDate(GridSaisie.Cells[1,ARow]), Periode* StrToInt(GridSaisie.Cells[3,ARow]))-1));
       4,5 : begin
              if IsFloat (Cellule) or IsNumeric (Cellule) then
               GridSaisie.Cells[Acol,ARow] := StrfMontant ( Valeur(Cellule),20, V_PGI.OkDecV , '', false)
               else Cancel := true;
             end;
     end;
   end
   else
   begin
     case ACol of
       1: if ARow = 1 then Cancel := true;
       3: begin
          if ((GridSaisie.Cells[2,ARow]<>'') and (GridSaisie.Cells[3,ARow]<>'')) then
            GridSaisie.Cells[3,ARow] := IntToStr (CalculNombreEcheance (StrToDate(GridSaisie.Cells[1,ARow]),
                                                 StrToDate(GridSaisie.Cells[2,ARow]), Periode));
          end;
      end;
   end;
   if (ACol = 3) and (ARow = GridSaisie.RowCount - 1) then
       GridSaisie.RowCount := GridSaisie.RowCount + 1;
end;

procedure TFPlanEche.GridSaisieRowExit(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
  if  (IsLigneVide (Ou)) or (IsLigneOK(Ou)) then
  begin
    Cancel := false;
    CalculCumul;
  end
  else Cancel := true;
end;

procedure TFPlanEche.CalculCumul;
var  nEcheance : integer;
     MontantLoyer, MontantFrais : double;
     i : integer;
begin
  nEcheance := 0;
  MontantLoyer := 0.0;
  MontantFrais := 0.0;
  for i := 1 to GridSaisie.RowCount do
  if IsLigneOK(i) then
  begin
    nEcheance := nEcheance + StrToInt(GridSaisie.Cells[3,i]);
    MontantLoyer := MontantLoyer + StrToInt(GridSaisie.Cells[3,i])*Valeur(GridSaisie.Cells[4,i]);
    MontantFrais := MontantFrais + StrToInt(GridSaisie.Cells[3,i])*Valeur(GridSaisie.Cells[5,i]);
  end;
  TotalNombre.Value := nEcheance;
  TotalValeur.Value := MontantLoyer;
  TotalFrais.Value := MontantFrais;
end;

procedure TFPlanEche.FormShow(Sender: TObject);
var sPeriode : string;
begin
     GridSaisie.ColAligns[1]:= taCenter;
     GridSaisie.ColAligns[2]:= taCenter;
     GridSaisie.ColAligns[3]:= taCenter;
     GridSaisie.ColAligns[4]:= taRightJustify;
     GridSaisie.ColAligns[5]:= taRightJustify;
     case Periode of
       1: sPeriode := 'MEN';
       3: sPeriode := 'TRI';
       6: sPeriode := 'SEM';
       12: sPeriode := 'ANN';
     end;
     PERIODICITE.Text := RechDom('TIFREQUENCE',sPeriode,FALSE);
     VERSEMENT.Text := RechDom('TITYPEVERSEMENT',fVersement,FALSE);
end;

procedure TFPlanEche.InitGridSaisie(Date : TDateTime;LTranche : TList);
var i:integer;
    ARecord : PTrancheEcheance;
begin
  if (LTranche <> nil) then
  begin
    GridSaisie.RowCount := LTranche.Count + 2;
    if LTranche.Count = 0 then GridSaisie.Cells[1,1] := DateToStr(Date);
    for i := 0 to LTranche.Count - 1 do
    begin
      ARecord :=  LTranche.Items[i];
      GridSaisie.Cells[1,i+1] := DateToStr(ARecord^.DateDebut);
      GridSaisie.Cells[2,i+1] := DateToStr(ARecord^.DateFin);
      GridSaisie.Cells[3,i+1] := IntToStr(ARecord^.nEcheance);
      GridSaisie.Cells[4,i+1] := StrfMontant(ARecord^.Montant,20,V_PGI.OkDecV,'',false);
      GridSaisie.Cells[5,i+1] := StrfMontant(ARecord^.Frais,20,V_PGI.OkDecV,'',false);
    end;
  end;
  CalculCumul;
end;


procedure TFPlanEche.EnregistreListeTrancheEcheance ( var LTranche : TList);
var i:integer;
    ARecord : PTrancheEcheance;
begin
  // Destruction de la liste de tranches d'échéances en cours
  if LTranche <> nil then
  begin
    for i := 0 to (LTranche.Count - 1) do
    begin
      ARecord := LTranche.Items[i];
      Dispose(ARecord);
    end;
    Ltranche.Free;
    Ltranche := nil;
  end;

  // Création et mise à jour de la liste de tranches d'échéances
  LTranche := TList.Create;
  for i := 1 to GridSaisie.RowCount do
  if GridSaisie.Cells[1,i] <> '' then
  begin
    new (ARecord);
    ARecord.DateDebut := StrToDate(GridSaisie.Cells[1,i]);
    ARecord.DateFin := StrToDate(GridSaisie.Cells[2,i]);
    if (GridSaisie.Cells[3,i] <> '') then
       ARecord.nEcheance := StrToInt(GridSaisie.Cells[3,i])
    else ARecord.nEcheance := 0;
    if (GridSaisie.Cells[4,i] <> '') then
    ARecord.Montant := Valeur(GridSaisie.Cells[4,i])
    else     ARecord.Montant := 0.0;
    if (GridSaisie.Cells[4,i] <> '') then
    ARecord.Frais := Valeur(GridSaisie.Cells[5,i])
    else ARecord.Frais := 0.0;
    LTranche.Add (ARecord);
  end;
end;

function TFPlanEche.IsLigneVide (Row : integer): boolean;
begin
if (GridSaisie.Cells[1,Row] = '') and (GridSaisie.Cells[2,Row] = '')
and (GridSaisie.Cells[3,Row] = '') and (GridSaisie.Cells[4,Row] = '')
and (GridSaisie.Cells[5,Row] = '')
then result := true
else result := false;
end;

function TFPlanEche.IsLigneOK (Row : integer): boolean;
begin
if (GridSaisie.Cells[1,Row] <> '') and (GridSaisie.Cells[2,Row] <> '')
and (GridSaisie.Cells[3,Row] <> '') and (GridSaisie.Cells[4,Row] <> '')
and (GridSaisie.Cells[5,Row] <> '')
then
{begin
  if IsValidDate (GridSaisie.Cells[1,Row+1]) and (Row > 0) then
    if StrToDate (GridSaisie.Cells[1,Row+1]) <= StrToDate (GridSaisie.Cells[2,Row]) then
    begin
      Result := false;
       HM1.Execute (1,Caption,'');
      exit;
    end;
  if (Row >= 2) and IsValidDate (GridSaisie.Cells[2,Row-1]) then
    if StrToDate (GridSaisie.Cells[2,Row-1]) >= StrToDate (GridSaisie.Cells[1,Row]) then
    begin
      Result := false;
      exit;
    end;}
  result := true{;
end}
else result := false;
end;

function TFPlanEche.ControleCoherenceDate : boolean;
var i : integer;
begin
  for i:=1 to GridSaisie.RowCount - 1 do
  begin
    if IsLigneVide(i) then continue;
    if ((IsValidDate(GridSaisie.Cells[E_FIN,i])) and (IsValidDate(GridSaisie.Cells[E_DEBUT,i]))) then
    begin
      if StrToDate (GridSaisie.Cells[E_FIN,i]) <= StrToDate (GridSaisie.Cells[E_DEBUT,i]) then
      begin
        HM1.Execute (0,Caption,'');
        Result := false;
        GridSaisie.Row := i;
        exit;
      end;
    end else
    begin
        HM1.Execute (0,Caption,'');
        Result := false;
        GridSaisie.Row := i;
        exit;
    end;
    if i <= 1 then continue;
    if StrToDate (GridSaisie.Cells[E_DEBUT,i]) <= StrToDate (GridSaisie.Cells[E_FIN,i-1]) then
    begin
      HM1.Execute (2,Caption,'');
      Result := false;
      GridSaisie.Row := i;
      exit;
    end;
  end;
  Result := true;
end;

procedure TFPlanEche.GridSaisieColExit(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
    if Chg = true then exit;
    case Ou of
    1: if (GridSaisie.Row > 1) and (GridSaisie.Cells[E_DEBUT,GridSaisie.Row]='') then
       GridSaisie.Cells[E_DEBUT,GridSaisie.Row] := DateToStr(StrToDate (GridSaisie.Cells[E_FIN,GridSaisie.Row-1]) + 1);
    end;
end;

procedure TFPlanEche.GridSaisieRowEnter(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
     if not (IsLigneOk (Ou-1)) then Cancel := true;
end;

procedure TFPlanEche.GridSaisieGetEditMask(Sender: TObject; ACol,
  ARow: Integer; var Value: String);
begin
//  if (ACol = 1) or (ACol = 2) then Value := '##/##/####'; car pb effacement
end;

procedure TFPlanEche.BValideClick(Sender: TObject);
var i : integer;
begin
  for i := 1 to GridSaisie.RowCount - 1 do
  begin
    if not IsLigneVide (i) then
    begin
      if not IsLigneOK (i) then
      begin
        ModalResult := mrNone;
        HM1.Execute(3,Caption,'');
        GridSaisie.Row := i;
        exit;
      end;
    end
    else
    begin
      if i < GridSaisie.RowCount-1 then
      begin
        HM1.Execute(4,Caption,'');
        GridSaisie.Row := i;
        ModalResult := mrNone;
        exit;
      end;
    end;
  end;
  if not ControleCoherenceDate then
  begin
    ModalResult := mrNone;
    exit;
  end;
  ModalResult := mrOK;
end;

procedure TFPlanEche.BAbandonClick(Sender: TObject);
var mr : integer;
begin
  mr := HM1.Execute (5,Caption,'');
  if mr = mrYes then ModalResult := mrNo else ModalResult := mrNone;
end;

procedure TFPlanEche.BImprimerClick(Sender: TObject);
begin
{$IFDEF eAGLCLient}
//YCP PrintDBGrid(Caption,Q.Liste,Q.SQL.Text,Q.UserListe) ;
{$ELSE}
   PrintDBGrid (GridSaisie,Nil, Caption,'');
{$ENDIF}

end;

procedure TFPlanEche.InsererClick(Sender: TObject);
var i,j : integer;
begin
  if IsLigneOK (GridSaisie.Row) then
  begin
    GridSaisie.RowCount := GridSaisie.RowCount + 1;
    for i := GridSaisie.RowCount downto GridSaisie.Row+1 do
      for j := 1 to 5 do
      begin
        GridSaisie.Cells[j,i]:= GridSaisie.Cells[j,i-1];
        GridSaisie.Cells[j,i-1] := '';
      end;
    end;
end;

procedure TFPlanEche.SupprimerClick(Sender: TObject);
var i,j : integer;
begin
  for i:= GridSaisie.Row to GridSaisie.RowCount-1 do
    for j:= 1 to 5 do
    begin
      GridSaisie.Cells[j,i] :=  GridSaisie.Cells[j,i+1];
    end;
  if GridSaisie.RowCount > 2 then GridSaisie.RowCount := GridSaisie.RowCount - 1
  else
  for i := 1 to 5 do GridSaisie.Cells[i,1]:='';
  CalculCumul;
end;

procedure TFPlanEche.HelpBtnClick(Sender: TObject);
begin
CallHelpTopic(Self);
end;

end.
