{***********UNITE*************************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 20/06/2007
Modifié le ... : 11/09/2007
Description .. : Saisie des échéances en cas de contrat à loyer variable
Suite ........ : 09/2007 Suite FQ 20005 Bloquer la colonne des dates
Mots clefs ... : 
*****************************************************************}
unit planecheeche;

interface

uses
  SysUtils, Classes, Controls, Forms,
  Grids, Hctrls, HPanel, StdCtrls, HEnt1, hmsgbox,
  HTB97, HSysMenu, Menus, ExtCtrls;

function  FicheSaisieEcheancierEche (DateTranche : TDateTime;Periode : integer; Versement:string; LEcheances : TList) : boolean;

const ECH_DEBUT=1;
      ECH_MONTANT=2; //4;
      ECH_FRAIS=3;  //5;

type
  PEche = ^AEche;
  AEche = record
    Date: TDateTime;
    Montant,Frais:double;
    Echu : boolean;   // échéance cumulée non visible si True
  end;
  TFPlanEcheEche = class(TForm)
    HM1: THMsgBox;
    HPanel3: THPanel;
    HLabel3: THLabel;
    TotalValeur: THNumEdit;
    TotalFrais: THNumEdit;
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
    TotalValZ: THNumEdit;
    TotalFraisZ: THNumEdit;
    sDateZ: THLabel;
    HLabel4: THLabel;
    procedure GridSaisieCellExit(Sender: TObject; var ACol, ARow: Integer;
      var Cancel: Boolean);
    procedure FormShow(Sender: TObject);
    procedure GridSaisieRowExit(Sender: TObject; Ou: Integer;
      var Cancel: Boolean; Chg: Boolean);
    procedure GridSaisieCellEnter(Sender: TObject; var ACol, ARow: Integer;
      var Cancel: Boolean);
    procedure GridSaisieColExit(Sender: TObject; Ou: Integer;
      var Cancel: Boolean; Chg: Boolean);
    procedure GridSaisieRowEnter(Sender: TObject; ARow: Integer;
      var Cancel: Boolean; Chg: Boolean);
    procedure BValideClick(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure HelpBtnClick(Sender: TObject);
  private
    { Déclarations privées }
    Periode : integer;
    stCur : string;
    fVersement : string;
    fLInit : TList;  // Echéancier copié en entrée à partir de LEcheances
    fLEche : Tlist;  // Echéancier extrait de fLInit et chargé par le THGrid
    fDateZ : TDateTime;
    fRienEchu : Boolean;   // False <=> présence échéances cumulées
    DateAchat : TDateTime;
    procedure EnregistreListeEcheance ( LEcheances : TList);
    procedure InitGridSaisie(Date : TDateTime; LEcheances : TList);
    procedure CalculCumul;
    //function IsLigneVide (Row : integer): boolean;
    function IsLigneOK (Row : integer): boolean;
    //function ControleCoherenceDate : boolean; 
    procedure CopieEcheancierInit (LEcheances : TList);
    procedure ExtraireEcheancierEncours;
    procedure DetruitListes;
  public
    { Déclarations publiques }
  end;


implementation

uses Outils, ImEnt, ImContra
{$IFDEF eAGLCLient}
   ;
{$ELSE}
   , PrintDBG;
{$ENDIF}

{$R *.DFM}

function  FicheSaisieEcheancierEche (DateTranche : TDateTime;Periode : integer; Versement:string;LEcheances : TList) : boolean;
var
  FicheEcheancier: TFPlanEcheEche;
begin
FicheEcheancier := TFPlanEcheEche.Create(Application);
FicheEcheancier.Periode := Periode;
FicheEcheancier.DateAchat := DateTranche;
FicheEcheancier.fVersement := Versement;
try
FicheEcheancier.InitGridSaisie(DateTranche, LEcheances);
if FicheEcheancier.ShowModal =  mrOk then
begin
  Result := True;
  FicheEcheancier.EnregistreListeEcheance (LEcheances);
end else Result := False;
finally
FicheEcheancier.DetruitListes;
FicheEcheancier.GridSaisie.VidePile (True);
FicheEcheancier.Free;
end;
end;

procedure TFPlanEcheEche.GridSaisieCellEnter(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
  stCur := GridSaisie.Cells [ACol,Arow];
end;

procedure TFPlanEcheEche.GridSaisieCellExit(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
var Cellule : string;
begin
   Cellule := GridSaisie.Cells[ACol,ARow];
   if Cellule <> '' then
   begin
     if (Cellule <> stCur) then
     case ACol of
{      1 : begin
               if IsValidDate(Cellule) = false then
                  Cancel := true
               else
               begin
                 GridSaisie.Cells[ACol,ARow] := DateToStr(StrToDate(Cellule));
                 // 1e échéance se compare à la dernière date échue si existe
                 // sinon à la date de livraison
                 if (ARow = 1) and (not fRienEchu) then
                    if (StrToDate (Cellule) <= StrToDate(sDateZ.Caption)) then
                        HM1.Execute(0, Caption, '');
                 if (ARow = 1) and (fRienEchu) then
                    if (StrToDate (Cellule) < DateAchat) then
                        HM1.Execute(0, Caption, '');
                 if ARow > 1 then
                    if (StrToDate (Cellule) <= StrToDate(GridSaisie.Cells[1,ARow-1])) then
                        HM1.Execute(0, Caption, '');
               end;
           end;    }
       2,3 : begin
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
      end;
   end;
   //if (ACol = 2) and (ARow = GridSaisie.RowCount - 1) then
   //    GridSaisie.RowCount := GridSaisie.RowCount + 1;
end;

procedure TFPlanEcheEche.GridSaisieRowExit(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
  //if  (IsLigneVide (Ou)) or (IsLigneOK(Ou)) then
  if (IsLigneOK(Ou)) then
  begin
    Cancel := false;
    CalculCumul;
  end
  else Cancel := true;
end;

procedure TFPlanEcheEche.CalculCumul;
var  MontantLoyer, MontantFrais : double;
     i : integer;
begin
  MontantLoyer := 0.0;
  MontantFrais := 0.0;
  for i := 1 to GridSaisie.RowCount do
     if IsLigneOK(i) then
     begin
      MontantLoyer := MontantLoyer + Valeur(GridSaisie.Cells[2,i]);
      MontantFrais := MontantFrais + Valeur(GridSaisie.Cells[3,i]);
     end;
  TotalValeur.Value := MontantLoyer + TotalValZ.Value;
  TotalFrais.Value := MontantFrais + TotalFraisZ.Value;
end;

procedure TFPlanEcheEche.FormShow(Sender: TObject);
var sPeriode : string;
begin
     //GridSaisie.ColEditables[1]:=False;
     GridSaisie.ColAligns[1]:= taCenter;
     GridSaisie.ColAligns[2]:= taRightJustify;
     GridSaisie.ColAligns[3]:= taRightJustify;

     case Periode of
       1: sPeriode := 'MEN';
       3: sPeriode := 'TRI';
       6: sPeriode := 'SEM';
       12: sPeriode := 'ANN';
     end;
     PERIODICITE.Text := RechDom('TIFREQUENCE',sPeriode,FALSE);
     VERSEMENT.Text := RechDom('TITYPEVERSEMENT',fVersement,FALSE);
end;

procedure TFPlanEcheEche.InitGridSaisie(Date : TDateTime;LEcheances : TList);
var i :integer;
    ARecord : PEche;
begin

  // Recopier LEcheances dans fLInit
  CopieEcheancierInit (LEcheances);

  // Remplir fLEche avec la partie de fLInit des échéances non échues
  ExtraireEcheancierEncours;

  if (fLEche <> nil) then
  begin
    GridSaisie.RowCount := fLEche.Count + 1;   //2;

    for i := 0 to fLEche.Count - 1 do
    begin
      ARecord :=  fLEche.Items[i];
      GridSaisie.Cells[1,i+1] := DateToStr(ARecord^.Date);
      GridSaisie.Cells[2,i+1] := StrfMontant(ARecord^.Montant,20,V_PGI.OkDecV,'',false);
      GridSaisie.Cells[3,i+1] := StrfMontant(ARecord^.Frais,20,V_PGI.OkDecV,'',false);
    end;

    // Echéancier non échu est vide, ne rien saisir
    if fLEche.Count = 0 then
    begin
       GridSaisie.Options := GridSaisie.Options - [goEditing] - [goAlwaysShowEditor];
       GridSaisie.RowCount := 2;
       GridSaisie.Row := 1;
       //if fRienEchu then
       //   GridSaisie.Cells[1,1] := DateToStr(Date)
       //else
       //   GridSaisie.Cells[1,1] := DateToStr(PlusMois (StrToDate(sDateZ.Caption),Periode));
    end;
  end;
  CalculCumul;
end;


procedure TFPlanEcheEche.EnregistreListeEcheance ( LEcheances : TList);
var i:integer;
    ARecord:PEcheance;
    BRecord:PEche;
begin
  // Destruction de la liste des échéances donnée en entrée
  if LEcheances <> nil then
  begin
    for i := 0 to (LEcheances.Count - 1) do
    begin
      ARecord := LEcheances.Items[i];
      Dispose(ARecord);
    end;
    LEcheances.Clear;
  end;

  // Echéancier échu caché à enregistrer en premier
  if (fLInit <> nil) then
  begin
     for i := 0 to (fLInit.Count - 1) do
     begin
       BRecord := fLInit.Items[i];
       if BRecord^.Echu  then
       begin
          new (ARecord);
          ARecord^.Date := BRecord^.Date;
          ARecord^.Montant := BRecord^.Montant;
          ARecord^.Frais := BRecord^.frais;
          LEcheances.Add (ARecord);
       end;
     end;
  end;

  // Echéancier saisi à enregistrer ensuite
  for i := 1 to GridSaisie.RowCount do
    if GridSaisie.Cells[1,i] <> '' then
    begin
      new (ARecord);
      ARecord.Date := StrToDate(GridSaisie.Cells[1,i]);
      if (GridSaisie.Cells[2,i] <> '') then
         ARecord.Montant := Valeur(GridSaisie.Cells[2,i])
      else     ARecord.Montant := 0.0;
      if (GridSaisie.Cells[2,i] <> '') then
         ARecord.Frais := Valeur(GridSaisie.Cells[3,i])
      else ARecord.Frais := 0.0;
      LEcheances.Add (ARecord);
    end;
end;

{function TFPlanEcheEche.IsLigneVide (Row : integer): boolean;
begin
if (GridSaisie.Cells[1,Row] = '')
and (GridSaisie.Cells[2,Row] = '')
and (GridSaisie.Cells[3,Row] = '')
then result := true
else result := false;
end; }

function TFPlanEcheEche.IsLigneOK (Row : integer): boolean;
begin

if Row > fLEche.count then
begin
   GridSaisie.Cells[ECH_MONTANT,Row] := '';
   GridSaisie.Cells[ECH_FRAIS,Row] := '';
end;

//if (GridSaisie.Cells[1,Row] <> '')
if  (GridSaisie.Cells[2,Row] <> '')
and (GridSaisie.Cells[3,Row] <> '')
then result := true
else result := false;
end;

{function TFPlanEcheEche.ControleCoherenceDate : boolean;
var i : integer;
begin
  for i:=1 to GridSaisie.RowCount - 1 do
  begin
    if IsLigneVide(i) then continue;
    if not (IsValidDate(GridSaisie.Cells[ECH_DEBUT,i])) then
    begin
        HM1.Execute (0,Caption,'');
        Result := false;
        GridSaisie.Row := i;
        exit;
    end;
    if i < 1 then continue;

    // 1e échéance à contrôler avec la dernière échue si existe
    if i=1 then
    begin
       if (not fRienEchu) then
       begin
          if (StrToDate (GridSaisie.Cells[ECH_DEBUT,i]) <> PlusMois (StrToDate(sDateZ.Caption),Periode)) then
          begin
            HM1.Execute(0, Caption, '');
            Result := false;
            GridSaisie.Row := i;
            exit;
          end;
       end;

    end else
    begin
       if (StrToDate (GridSaisie.Cells[ECH_DEBUT,i]) <> PlusMois (StrToDate(GridSaisie.Cells[ECH_DEBUT,i-1]),Periode)) then
       begin
          HM1.Execute(2, Caption, '');
          Result := false;
          GridSaisie.Row := i;
          exit;
       end;
    end;
  end;
  Result := true;
end; }

procedure TFPlanEcheEche.GridSaisieColExit(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
     if Chg = true then exit;
end;

procedure TFPlanEcheEche.GridSaisieRowEnter(Sender: TObject; ARow: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
  if not (IsLigneOk (ARow-1)) then
     begin
     Cancel := true;
     exit;
     end;

{  if GridSaisie.Cells[ECH_MONTANT,ARow] <> '' then exit;

 if (IsFloat(GridSaisie.Cells[ECH_MONTANT,ARow-1])) or
           (IsNumeric(GridSaisie.Cells[ECH_MONTANT,ARow-1])) then
                         GridSaisie.Cells[ECH_MONTANT,ARow]:= GridSaisie.Cells[ECH_MONTANT,ARow-1]
  else GridSaisie.Cells[ECH_MONTANT,ARow]:= MontantToStr(0.00);

  if GridSaisie.Cells[ECH_FRAIS,ARow] <> '' then exit;

 if (IsFloat(GridSaisie.Cells[ECH_FRAIS,ARow-1])) or
           (IsNumeric(GridSaisie.Cells[ECH_FRAIS,ARow-1])) then
                  GridSaisie.Cells[ECH_FRAIS,ARow]:= GridSaisie.Cells[ECH_FRAIS,ARow-1]
  else GridSaisie.Cells[ECH_FRAIS,ARow]:= MontantToStr(0.00); }

  if ARow > fLEche.count then
  begin
     GridSaisie.Cells[ECH_MONTANT,ARow] := '';
     GridSaisie.Cells[ECH_FRAIS,ARow] := '';
  end;
end;

procedure TFPlanEcheEche.BValideClick(Sender: TObject);
var i : integer;
begin
  for i := 1 to fLEche.Count do  //GridSaisie.RowCount - 1 do
  begin
    //if not IsLigneVide (i) then
    //begin

      if not IsLigneOK (i) then
      begin
        ModalResult := mrNone;
        HM1.Execute(3,Caption,'');
        GridSaisie.Row := i;
        exit;
      end;
    {end
    else
    begin
      if i < GridSaisie.RowCount-1 then
      begin
        HM1.Execute(4,Caption,'');
        GridSaisie.Row := i;
        ModalResult := mrNone;
        exit;
      end;
    end; }
  end;
{  if not ControleCoherenceDate then
  begin
    ModalResult := mrNone;
    exit;
  end; }
  ModalResult := mrOK;
end;

procedure TFPlanEcheEche.BAbandonClick(Sender: TObject);
var mr : integer;
begin
  mr := HM1.Execute (5,Caption,'');
  if mr = mrYes then ModalResult := mrNo else ModalResult := mrNone;
end;

procedure TFPlanEcheEche.BImprimerClick(Sender: TObject);
begin
{$IFDEF eAGLCLient}
{$ELSE}
   PrintDBGrid (GridSaisie,Nil, Caption,'');
{$ENDIF}
end;

procedure TFPlanEcheEche.HelpBtnClick(Sender: TObject);
begin
CallHelpTopic(Self);
end;


procedure TFPlanEcheEche.CopieEcheancierInit (LEcheances : TList);
var
  i:integer;
  wRecord : PEche;
  ARecord : PEcheance;
begin
  fLInit := TList.Create;

  if LEcheances <> nil then
  begin
    for i := 0 to (LEcheances.Count - 1) do
    begin
      ARecord := LEcheances.Items[i];
      New (wRecord);
      wRecord^.Date := ARecord^.Date;
      wRecord^.Montant := ARecord^.Montant;
      wRecord^.Frais := ARecord^.Frais;
      wRecord^.Echu := False;
      fLInit.Add (wRecord);
    end;
  end;

end;

// On ne garde de fLInit que les échéances modifiables, on synthétise les autres dans 'Total au'
// Echéances modifiables = celles commençant à l'exercice en cours
procedure TFPlanEcheEche.ExtraireEcheancierEncours;
var
  i:integer;
  ARecordInit, wRecord : PEche;
  garder : boolean;

begin
  fRienEchu:= True;
  fLEche := TList.Create;

  if fLInit <> nil then
  begin
    for i := 0 to (fLInit.Count - 1) do
    begin
      ARecordInit := fLInit.Items[i];

      // Versement à l'avance, on n'affiche que les échéances débutant à partir du début d'exercice
      // On n'affiche pas les échéances à cheval car déjà payées
      if (fVersement ='AVA') then
      begin
        if ARecordInit^.Date >= VHImmo^.Encours.Deb then
          garder := True
        else
          // antérieure ou à cheval
          garder := False;

      // Versement à terme échu, même chose
      end else
      begin
        if PlusMois((ARecordInit^.Date + 1), (-1)* Periode) >= VHImmo^.Encours.Deb then
          garder := True
        else
          // antérieure ou à cheval
          garder := False;
      end;

      // Echéance à afficher à l'écran
      if garder then
      begin
          ARecordInit^.Echu :=  False;
          new (wRecord);
          wRecord := fLInit.Items[i];
          fLEche.Add (wRecord);
      end else
      begin
      // Cumuler dans les totaux
          ARecordInit^.Echu :=  True;
          TotalValZ.Value := TotalValZ.Value +  ARecordInit^.Montant;
          TotalFraisZ.Value := TotalFraisZ.Value + ARecordInit^.Frais;
          fRienEchu := False;
          if ARecordInit^.Date > fDateZ then  fDateZ := ARecordInit^.Date;
      end; // garder
    end; // for
  end;   // fLInit

  // Si toutes les échéances sont affichées => Cacher les totaux, sinon afficher la dernière date échue
  if fRienEchu then
  begin
       sDateZ.Visible := False;
       TotalValZ.Visible := False;
       TotalFraisZ.Visible := False;
       HLabel4.Visible := False;
  end else
  begin
       sDateZ.Caption  := DateToStr (fDateZ);
  end;
end;

procedure TFPlanEcheEche.DetruitListes;
var i : integer;
begin
  // Attention : bien faire un .Delete et pas un Dispose du Record
  // car Dispose écrase la 2e TList fLEche => violation d'accès en pointeur incorrect
  if fLInit <> nil then
  begin
    for i:=(fLInit.count - 1) downto 0 do
      fLInit.Delete(i);
   fLInit.Free;
   fLInit := nil;
  end;
  if fLEche <> nil then
  begin
    for i:=(fLEche.count - 1) downto 0 do
      fLEche.Delete(i);
    fLEche.Free;
    fLEche := nil;
  end;

end;


end.
