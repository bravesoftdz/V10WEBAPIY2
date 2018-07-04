unit AFCalendrier;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HSysMenu, AGLInit, UIUtil, HPanel, HCtrls, Hent1, HTB97, StdCtrls,
  ExtCtrls, Mask, Buttons,
{$IFDEF EAGLCLIENT}
  eTablette,
{$ELSE}
  Tablette,
{$ENDIF}
  hmsgbox, Grids,UAFO_Ferie;

  Type TDEFRECT = RECORD
        NbLigneInRect   : integer;
        NbColInRect     : integer;
        IsBold          : Boolean;
        Cadrage         : TAlignment;
        End;

     // Classe du composant Graphique Calendrier
 Type TGridCalendrier = Class
    Private
    FMonthOffset: Integer;
    FStartOfWeek: Integer; // Jour de début du calendrier
    FNumSemaine : Boolean; // Affichage des numéros de semaine
    FJourFermeture : array[0..6] of Integer;    // Liste des jours de fermeture
    FDecal, FNbRow, FNbJourSemaine, FNbColNonFixe : Integer;
    FGestionFerie : Boolean;
    FJourFerie : TJourFerie;
    FLabelMois : TLabel;
    FCumul : Boolean;
    OldBrush : TBrush;
    OldPen : TPen;
    Procedure LibelleMois;
    Procedure CalendrierDrawCell(Sender: TObject; Col, Row: Longint;
                 Rect: TRect; State: TGridDrawState);

    Protected
    Ecran : TForm;
    FDefRect : TDefRect;
    FGriseJourFermeture : Boolean;
    FDateDeb,FDateFin : TDateTime;
    // Fonction de chargement des objets à associer au Grid sur la période
    Procedure ChargeSpecifGridCalendrier; virtual;
    // Affectation des objets à chaque Jour du grid.
    Procedure ObjetCalendrierJour (LeJour:TDateTime; Col,Row,NumSemaine:integer); virtual;
    // Dessin pour chaque cellule des zones complémentaires au numéro de jour (géré en standard)
    Procedure CalendrierDrawRect(Col,Row :integer;Rect:TRect;Ferie:boolean); virtual;
    procedure CalendrierSelectCell(Sender: TObject; Col, Row: Longint; var CanSelect: Boolean); virtual;
    procedure CalendrierDblClick(Sender: TObject); virtual;
    procedure CalendrierCellEnter(Sender: TObject; var ACol, ARow: Integer;  var Cancel: Boolean); virtual;

    Public
    GrCal : THGrid;
    FColDebJour : integer;  // colonne de début des jours
    FYear,FMonth,FDay : word ;
    FCalendMois : Boolean; // Calendrier mensuel (mois par mois) ou perpétuel (déroulant)
    property NbRow  : Integer Read FNbRow;
    property NbJourSemaine  : Integer Read FNbJourSemaine;
    property JourFerie : TJourFerie Read FJourFerie;
    Function DateSelect(DateDeb:TDateTime;Col,Row:Integer):TDateTime;
    Procedure CelluleDunJour(DateCherchee : TDateTime; var ACol,ARow : integer);
    Procedure ModifJour(DateCible : TDateTime);
    Procedure ChangerDate(DateCible : TDateTime; bForcer:boolean);
    Procedure ChargeCalendrier;
    Constructor Create (GridCalen : THGrid; Form : TForm; Start,Row : Integer; Mensuel,
                AffNumSemaine,AffCumulsemaine, Ferie : Boolean; Date : TDateTime; LibMois:TLabel; DefRect : TDefRect);
    Destructor Destroy; override;
    Function  PremierJourAffiche(DateEnc : TDateTime) : TDateTime;
    Procedure VideObjetsCalendrier;
    End;

Const cllightGray = $00EEEEEE;

implementation
//***************** Fonction de la class GridCalendrier  ***********************

Constructor TGridCalendrier.Create (GridCalen : THGrid; Form : TForm;Start,Row : Integer;
                Mensuel, AffNumSemaine,AffCumulsemaine,Ferie : Boolean; Date : TDateTime; LibMois:TLabel; DefRect : TDefRect);
var i,AColRow,LngCol0,LngColCumul : integer;
Begin
LngCol0:=0;
// Paramétrage du type de calendrier
// Commencer le calendrier au ... lundi = 1 etc ....
If ((Start > 0) And (Start <=7)) Then FStartOfWeek:=Start Else FStartOfWeek := 1;
// Nombre de lignes (semaines) dans le calendrier
If ((Row >0) And (Row <20)) Then FNbRow := Row Else FNbRow := 6;

FCalendMois:=Mensuel;        // Calendrier mensuel (non déroulant)
FNumSemaine := AffNumSemaine;// Affichage des num de semaine en colonne 0
GrCal := GridCalen;          // Grid Calendrier
FGestionFerie :=Ferie;        // Gestion des jours Feriés
FCumul := AffCumulsemaine;   // Gestion d'une colonne de cumul par semaine
FDefRect:= DefRect;
FNbJourSemaine := 7;        // Nombres de jours affichés par semaine
// conservé du grid passé en argument :  DefaultRowHeight DefaultColWidth

// Libellé utilisé pour affichage du mois en cours
If (LibMois Is TLabel) Then FLabelMois := LibMois Else FLabelMois:= Nil;
Ecran := Form;   // Nécessaire pour l'accès aux composants du Form des méthodes surchargeables

OldBrush:= TBrush.Create;
OldPen:= TPen.Create;

FGriseJourFermeture := True;
// Evénements du Grid
GrCal.OnDrawCell:=CalendrierDrawCell;
GrCal.OnCellEnter:= CalendrierCellEnter;
GrCal.OnDblClick:= CalendrierDblClick;
GrCal.OnSelectCell:= CalendrierSelectCell;

// Définition du grid
GrCal.RowCount:= FNbRow+1;
GrCal.RowHeights[0]:=Trunc(GridCalen.Font.Size*GridCalen.Font.PixelsPerInch/72 + 4);
GrCal.Height := GrCal.RowHeights[0]+ (FNbRow * GrCal.DefaultRowHeight)+ (2*FNbRow);

GrCal.FixedCols:=0;
GrCal.ColCount := FNbJourSemaine; FNbColNonFixe := FNbJourSemaine;
FColDebJour:=0; FDecal:=1;

if FCumul then BEGIN GrCal.ColCount:=GrCal.ColCount + 1;  Inc(FNbColNonFixe); END;

If (FNumSemaine) then
   Begin
   GrCal.ColCount:=GrCal.ColCount + 1;
   GrCal.FixedCols:=1; GrCal.ColWidths[0]:=15;
   LngCol0:= GrCal.ColWidths[0]+ GrCal.GridLineWidth;
   FColDebJour:=1; FDecal:=0;
   end;


If(FCalendMois) Then
    Begin
    GrCal.ScrollBars:=ssNone;
    LngColCumul := 0;
    if FCumul then
        BEGIN
        GrCal.ColWidths[GrCal.Colcount-1]:= Trunc(GrCal.DefaultColWidth/2)+3;
        LngColCumul := Trunc(GrCal.DefaultColWidth/2) + GrCal.GridLineWidth+3;
        END;
    GrCal.Width := ((GrCal.DefaultColWidth+GrCal.GridLineWidth)*FNbJourSemaine)+LngCol0+GrCal.GridLineWidth+ 4 + LngColCumul;
    End;
//                Else Begin GrCal.ScrollBars:=ssVertical; GrCal.Width := (GrCal.DefaultColWidth+GrCal.GridLineWidth)*7+10; End;

// Affichage des libellés sur la ligne fixe
For AColRow:=FColDebJour to (FNbJourSemaine + FColDebJour -1) do
    Begin
    i:=(FStartOfWeek + AColRow)mod 7 + FDecal;
    if (i=0) then i:=7;
    GrCal.Cells[AColRow,0]:=Copy(ShortDayNames[i],1,3);
    End;
if FCumul then GrCal.Cells[GrCal.ColCount-1,0]:='Total';

GRCal.Canvas.Font.Name:=GrCal.Font.Name; GRCal.Canvas.Font.Size:=GrCal.Font.Size;
{$IFDEF EAGLCLIENT}
GrCal.Invalidate ; 
{$ELSE}
GrCal.Repaint;
{$ENDIF}
DecodeDate(Date,FYear,FMonth,FDay);
// Chargement des jours fériés
//If (FGestionFerie) Then FJourFerie:=TJourFerie.Create Else begin FJourFerie.free; FJourFerie:=Nil;end;
If (FGestionFerie) Then FJourFerie:=TJourFerie.Create Else begin FJourFerie:=Nil;end;
// Récupération des jours de fermeture
If (FGriseJourFermeture) Then ChargeJourFermeture(FJourFermeture);
// Chargement du calendrier
ChargeCalendrier;
End;

destructor TGridCalendrier.Destroy;
begin
OldBrush.Free;
OldPen.Free;
FJourFerie.Free; FJourFerie:=nil;
if (GrCal<>nil) then
    GrCal.VidePile(True);
inherited;
end;

Procedure TGridCalendrier.ChargeCalendrier;
Var IDay,IMonth,IYear : Word;
    C  : TDateTime ;
    i,j : integer;
    eOnChange:TSelectCellEvent;
BEGIN
// test de cohérence
if DaysPerMonth(FYear,FMonth)<FDay then FDay:=DaysPerMonth(FYear,FMonth);
C:=EncodeDate(FYear,FMonth,FDay);
// Détermine la date de début et de fin du calendrier
If (Not FCalendMois) Then
   Begin
   FDateDeb := PlusDate (C,-1,'A'); // planning glissant sur 2 ans
   FDateFin := PlusDate (C,1,'A');
   end Else
   Begin
   FDateDeb := EncodeDate(FYear,FMonth,1);
   FDateFin := FDateDeb +(7*FNbRow);
   End;
// Date Début et fin calées sur les 1er / dernier jours de la semaines
FMonthOffset:=2-((DayOfWeek(FDateFin)-FStartOfWeek+7) mod 7);
if FMonthOffset = 2 then FMonthOffset := -5;
FDateFin := FDateFin + FMonthOffset + 5;
FMonthOffset:=2-((DayOfWeek(FDateDeb)-FStartOfWeek+7) mod 7); //day of week for 1st of month
if FMonthOffset = 2 then FMonthOffset := -5;
FDateDeb := FDateDeb + FMonthOffset -1; // date début correspondant au 1er jour de la semaine

If (Not FCalendMois) Then Begin FNbRow :=(Trunc(FDateFin+1 - FDateDeb) div 7)+1; GrCal.RowCount:=FNbRow; end;
GrCal.TopRow:=((Trunc(EncodeDate(FYear,FMonth,1)-FDateDeb)) div 7)+1;
i:=(trunc(C-FDateDeb)div 7)+1;
eOnChange := GrCal.OnSelectCell;
GrCal.OnSelectCell:= nil;
GrCal.Row := i;
GrCal.OnSelectCell:= eOnChange;
i:=(Trunc(C-FDateDeb)mod 7)+FColDebJour;
GrCal.Col :=i;
LibelleMois;
// Initialisation des cellules
For i:=0 to GrCal.ColCount-1 do
  For j:=1 to FNbRow do GrCal.Cells[i,j]:='';
GrCal.Update;
// Appel de la fonction surchargeable de chargement spécifique
ChargeSpecifGridCalendrier;

// Chargement du Grid
For i:=FColDebJour to GrCal.ColCount-1 do
  For j:=1 to FNbRow do
  BEGIN
//    If (GrCal.Objects[i,j]<> nil) Then
//       GrCal.Objects[i,j]:=nil;

    if (FCumul) And (i = GrCal.ColCount-1) then
        BEGIN
        // la cellule cumul peut avoir un objet mais pas de jour associé
        // PL : on passe le numéro de semaine à la place de la colonne dans ce cas
        ObjetCalendrierJour(0,i,j,NumSemaine(DateSelect(FDateDeb,i-1,j)));
        END
    else
        BEGIN
        DecodeDate(DateSelect(FDateDeb,i,j),IYear,IMonth,IDay);

        //If ((Not FCalendMois) And (IDay=1)) Then GrCal.Cells[i,j]:=IntToStr(IDay)+' '+FirstMajuscule(LongMonthNames[IMonth])
        GrCal.Cells[i,j]:=IntToStr(IDay);
        ObjetCalendrierJour(DateSelect(FDateDeb,i,j),i,j,NumSemaine(DateSelect(FDateDeb,i-1,j))); // fonction surchargeable
        END;
  END;

END;


Procedure TGridCalendrier.LibelleMois;
Var IDayDeb,IMonthdeb,IYearDeb,IDayFin,IMonthFin,IYearFin : Word ;
Begin
If ((FLabelMois <>Nil) And (FLabelMois Is TLabel)) Then
    Begin
    DecodeDate(FDateDeb+6+FColDebJour+(GrCal.TopRow-1)*7,IYearDeb,IMonthDeb,IDayDeb);
    DecodeDate(FDateDeb+6+FColDebJour+(GrCal.TopRow+2)*7,IYearFin,IMonthFin,IDayFin);
    If (FCalendMois) Then  Begin IYearFin:=IYearDeb ; IMonthFin:=IMonthDeb; End; // pas de période à cheval sur 2 mois

    If (IYearDeb <> IYearFin)then
    FLabelMois.Caption:= FirstMajuscule(LongMonthNames[IMonthDeb])+' '+IntTostr(IYearDeb)+'-'+FirstMajuscule(LongMonthNames[IMonthFin])+' '+IntTostr(IYearFin) else
    If (IMonthdeb <> IMonthFin)then
    FLabelMois.Caption:= FirstMajuscule(LongMonthNames[IMonthDeb])+'-'+FirstMajuscule(LongMonthNames[IMonthfin])+' '+IntTostr(IYearFin) else
    FLabelMois.Caption:= FirstMajuscule(LongMonthNames[IMonthDeb])+' '+IntTostr(IYearFin);
    End;
End;

Function TGridCalendrier.DateSelect(DateDeb:TDateTime;Col,Row:Integer) :TDateTime;
Begin
Result :=DateDeb+(Col-FColDebJour)+(Row-1)*7;
End;

// Renvoie la colonne et la ligne du le calendrier correspondants à la date fournie en entrée :  DateCherchee
Procedure  TGridCalendrier.CelluleDunJour(DateCherchee : TDateTime; var ACol,ARow : integer);
Var     Ecart,NbSemaine,NbJour : integer;
Begin
Ecart := trunc(DateCherchee - FDateDeb);
// l'écart ne peut pas être négatif et supérieur à 6*7=42 (6 lignes de 7 jours affichés dans le calendrier)
If (Ecart < 0) Or (Ecart >42) Then Exit;
NbSemaine := Ecart div 7;
NbJour := Ecart - (NbSemaine*7);
ARow := NbSemaine+1;
ACol := NbJour + FColDebJour;
End;

// La cellule en cours du calendrier devient la cellule affectee à la date fournie en entrée : DateCible
Procedure  TGridCalendrier.ModifJour(DateCible : TDateTime);
Var Col, Row : integer;
eOnChange:TSelectCellEvent;
Begin
CelluleDunJour(DateCible, Col, Row);

eOnChange := GrCal.OnSelectCell;
GrCal.OnSelectCell:= nil;
GrCal.Row := Row;
GrCal.OnSelectCell:= eOnChange;

GrCal.Col := Col;
End;


Procedure  TGridCalendrier.ChangerDate(DateCible : TDateTime; bForcer:boolean);
var
DateCell : TDateTime;
iJourCal, iMoisCal, iAnneeCal:word;
iJourCible, iMoisCible, iAnneeCible:word;
begin
DateCell := DateSelect(FDateDeb,GrCal.Col,GrCal.Row);
if Not bForcer then
   if (DateCible=DateCell)  then exit;

DecodeDate(DateCell, iAnneeCal, iMoisCal, iJourCal);
DecodeDate(DateCible, iAnneeCible, iMoisCible, iJourCible);

if (iMoisCible<>FMonth) or ( iAnneeCible<>FYear) or bForcer then
   begin
   FYear := iAnneeCible;
   FMonth := iMoisCible;
   FDay := iJourCible;
   ChargeCalendrier;
   end
else
   ModifJour(DateCible);
// PL le 28/01/02 pour éviter plantage aléatoire en arrivant sur l'écran
{$IFDEF EAGLCLIENT}
//THGrid(Self).Invalidate ;
{$ELSE}
//TStringGrid(self).repaint;
{$ENDIF}
// Fin PL le 28/01/02 
end;

// Evènements associés au Grid Calendrier
procedure TGridCalendrier.CalendrierDrawCell(Sender: TObject; Col, Row: Longint; Rect: TRect; State: TGridDrawState);
var TheText: string;
    R : TRect;
    IDay,IMonth,IYear:Word;
    ferie, CellTypeDate, CellTypeCumul : Boolean;

begin
Ferie := false;
with GrCal.Canvas do
  BEGIN
  OldBrush.Assign(Brush);
  OldPen.Assign(Pen);
  R:=Rect;
  GrCal.Canvas.Brush.Style:=bsSolid ;
  TheText:=GrCal.Cells[Col,Row];
  Font.Name:=GrCal.Font.Name; Font.Size:=GrCal.Font.Size;
  GrCal.DefaultDrawing:=FALSE;
  IDay:=0; IMonth:=0; IYear:=0;

  CellTypeDate := true; CellTypecumul := False;
  if (Col = 0) and FNumSemaine  then CellTypeDate := False;
  if (Row = 0) and (TheText<>'')then CellTypeDate := False;
  if (FCumul)  and (Col = GrCal.ColCount-1) then BEGIN CellTypeDate := False; CellTypeCumul := True; End;

  if (CellTypeDate) then
      Begin
      If (Not FCalendMois) Then LibelleMois;
      try
      DecodeDate(DateSelect(FDateDeb,Col,Row),IYear,IMonth,IDay);
      except
      // L'exception sur la date doit être traitée
      IDay:=0;
      end;
      END;

  if fsBold in Font.Style then Font.Style:=Font.Style-[fsBold];
  // Traitement de la première ligne des jours de la semaine
  if (Row=0) then
      begin
      Brush.Color:=clBtnFace;
      Font.Color:=clWindowText;
      TextRect(Rect, Rect.Left + (Rect.Right - Rect.Left - TextWidth(TheText)) div 2,
          Rect.Top + (Rect.Bottom - Rect.Top - TextHeight(TheText)) div 2, TheText);
      Brush.Assign(OldBrush); Pen.Assign(OldPen); Exit;
      end ;
  // Traitement des numéros de semaine
  if ((Col =0) And (FNumSemaine)) Then
      Begin
      Brush.Color:=clBtnFace ;
      Font.Color:=clWindowText ;
      TheText:= IntToStr(NumSemaine(FDateDeb+(Row-1)*7));
      TextRect(Rect, Rect.Left + (Rect.Right - Rect.Left - TextWidth(TheText)) div 2,
          Rect.Top + (Rect.Bottom - Rect.Top - TextHeight(TheText)) div 2, TheText);
      Brush.Assign(OldBrush); Pen.Assign(OldPen); Exit;
      End;
  if (IDay = 0) And Not(CellTypeCumul) Then
      begin
      Brush.Color:=clLightGray ; Brush.Style:=bsSolid ;
      R.Left:=R.Left+1 ;R.Top:=R.Top+1 ;R.Right:=R.Right-1 ;R.Bottom:=R.Bottom-1 ;
      FillRect(R) ; Brush.Assign(OldBrush); Pen.Assign(OldPen); Exit ;
      end;

  Font.Color:=clWindowText ;

  if Not(CellTypeCumul) Then
        BEGIN
        // fond différent d'un mois sur l'autre
        If (FCalendMois) Then
          Begin
          If (IMonth <> FMonth) Then Begin Font.Color:=clSilver; Brush.Color:=clLightGray; End;
          End
        Else
          Begin
          If ((IMonth mod 2)=0) Then Brush.Color:=clLightGray
          else Brush.Color:=clWindow;
          End;

        // Traitement des jours de fermeture
        If ((FGriseJourFermeture) And
         (TestJourFermeture(DayOfWeek(EncodeDate(IYear,IMonth,IDay)),FJourFermeture))) then
         Font.Color:=clSilver;
        // Traitement des jours fériés
        If (FGestionFerie) And (FJourFerie.TestJourFerie(EncodeDate(IYear,IMonth,IDay))) then
        Begin Font.Color:=clRed; Ferie:=True End Else Ferie :=False;

        // Traitement de la cellule sélectionnée
        If ((gdSelected in State) or (gdFocused in State)) then
        begin
        Brush.Color:=clHighlight;
        if Font.Color=clWindowText then Font.Color:=clBtnHighlight ;
        end;

        // Affichage du jour
        If (GrCal.Objects[Col,Row]<> Nil) Then
        Begin
        R.Bottom := R.Top+ ((Rect.Bottom - Rect.Top) div FDefRect.NbLigneInRect);
        R.Right := R.Left +((Rect.Right - Rect.Left) div FDefRect.NbColInRect);
        End;
        If FDefRect.IsBold Then Font.Style:=Font.Style+[fsBold] ;
        If FDEfRect.Cadrage = taCenter Then
          TextRect(R, R.Left + (R.Right - R.Left - TextWidth(TheText)) div 2,
                  R.Top + (((Rect.Bottom - Rect.Top)div FDefRect.NbLigneInRect)- TextHeight(TheText))div 2, TheText)
        Else If FDEfRect.Cadrage = taLeftJustify Then
          TextRect(R, R.Left , R.Top + (((Rect.Bottom - Rect.Top)div FDefRect.NbLigneInRect)- TextHeight(TheText))div 2, TheText)
        Else If FDEfRect.Cadrage = taRightJustify Then
          TextRect(R, R.Right - TextWidth(TheText), R.Top + (((Rect.Bottom - Rect.Top)div FDefRect.NbLigneInRect)- TextHeight(TheText))div 2, TheText);


        If FDefRect.IsBold Then Font.Style:=Font.Style-[fsBold];
        END;

  CalendrierDrawRect(Col,Row,Rect,Ferie);  // Fonction surchargeable de dessin dans le rectangle.

   // encadre le jour actuel
  if Not(CellTypeCumul) Then
    If EncodeDate(IYear,IMonth,IDay)=Int(Now) then DrawEdge(Handle, Rect,EDGE_SUNKEN,BF_RECT);
  // Rechargement en Début / fin de calendrier déroulant
  If (Not FCalendMois) then
    Begin
    If (DateSelect(FDateDeb,Col,Row)=FDateFin) Then
       Begin
       DecodeDate(DateSelect(FDateDeb,Col,Row),FYear,FMonth,FDay);
       ChargeCalendrier;
       End;
    If (DateSelect(FDateDeb,Col,Row)=FDateDeb) Then
       Begin
       DecodeDate(DateSelect(FDateDeb,Col,Row),FYear,FMonth,FDay);
       ChargeCalendrier;
       End;
    End;
  Brush.Assign(OldBrush);
  Pen.Assign(OldPen);
  End;
end;

procedure TGridCalendrier.CalendrierSelectCell(Sender: TObject; Col, Row: Longint; var CanSelect: Boolean);
Var IDay,IMonth,IYear : Word ;
begin
If (Not FCalendMois) Then CanSelect:=(GrCal.Cells[Col,Row]<>'')
Else
    Begin
    DecodeDate(DateSelect(FDateDeb,Col,Row),IYear,IMonth,IDay);
    CanSelect := (IMonth = FMonth);
    if (FCumul) and (Col = GrCal.ColCount-1) then CanSelect := False;
    End;
end;

Procedure TGridCalendrier.ChargeSpecifGridCalendrier;
Begin
End;

Procedure TGridCalendrier.ObjetCalendrierJour (LeJour : TDateTime;Col,Row,NumSemaine :integer);
Begin
End;

procedure TGridCalendrier.CalendrierDblClick(Sender: TObject);
begin
end;

Procedure TGridCalendrier.CalendrierDrawRect(Col,Row : integer;Rect:TRect;Ferie:Boolean);
Begin
End;

procedure TGridCalendrier.CalendrierCellEnter(Sender: TObject; var ACol, ARow: Integer;  var Cancel: Boolean);
Begin
End;

Function  TGridCalendrier.PremierJourAffiche(DateEnc : TDateTime) : TDateTime;
Var Year,Month,Day : Word;
    MonthOffset : integer;
Begin
DecodeDate(DateEnc, Year, Month,Day );
result := encodeDate (Year,Month,1);
MonthOffset:=2-((DayOfWeek(Result)-FStartOfWeek+7) mod 7); //day of week for 1st of month
if MonthOffset = 2 then MonthOffset := -5;
Result := Result + MonthOffset -1; // date début correspondant au 1er jour de la semaine
End;

Procedure TGridCalendrier.VideObjetsCalendrier;
var
i,j : integer;
begin
// Chargement du Grid
For i:=FColDebJour to GrCal.ColCount-1 do
  For j:=1 to FNbRow do
       GrCal.Objects[i,j]:=nil;

end;

end.
