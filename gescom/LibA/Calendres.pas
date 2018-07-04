unit calendres;

interface

uses
  Windows, Messages, SysUtils, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Grids, Buttons, ExtCtrls, Hctrls, HEnt1, HSysMenu, HTB97, Mask,
  Menus, hmsgbox,
  ed_tools, HPanel,  UTob,ParamSoc,UAFO_Calendrier,AFCalendrier,Dicobtp, HeureUtil,
{$IFDEF EAGLCLIENT}
     UtileAGL,
{$ELSE}
     MajTable,HLines,DB, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$IFDEF V530}
     EdtEtat,
{$ELSE}
     EdtREtat,
{$ENDIF}
{$ENDIF}
     Classes ;
     
// Classe calendrier surchargée pour les ressources
Type  TGridCalendrierRessource = Class(TGridCalendrier)
    Procedure ChargeSpecifGridCalendrier; override;
    Procedure ObjetCalendrierJour (LeJour:TDateTime; Col,Row,NumSemaine:integer); override;
    Procedure CalendrierDrawRect(Col,Row :integer;Rect:TRect;Ferie:boolean); override;
    Procedure CalendrierDblClick(Sender: TObject); override;
    Procedure CalendrierCellEnter(Sender: TObject; var ACol, ARow: Integer;  var Cancel: Boolean); override;

    End;

type
  TFCalendrier = class(TForm)
    Cal: THGrid;
    LabelMois: TLabel;
    bNext: TToolbarButton97;
    bLast: TToolbarButton97;
    bPrevious: TToolbarButton97;
    bFirst: TToolbarButton97;
    FenHoraire: TToolWindow97;
    Hdeb1: THCritMaskEdit;
    Hfin1: THCritMaskEdit;
    HDeb2: THCritMaskEdit;
    HFin2: THCritMaskEdit;
    bValWindows: TToolbarButton97;
    bAnnulWindows: TToolbarButton97;
    HDuree: THCritMaskEdit;
    LabelStandard: THLabel;
    NomStandard: THLabel;
    CheckTravailFerie: TCheckBox;
    Dock971: TDock97;
    bValider: TToolbarButton97;
    bAnnuler: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    BImprimer: TToolbarButton97;
    HMTrad: THSystemMenu;
    ToolCalendrier: TToolWindow97;
    bDetail: TToolbarButton97;
    LDEB: THLabel;
    LDEB2: THLabel;
    LFIN1: THLabel;
    LFIN2: THLabel;
    LDUREE: THLabel;
    procedure FormShow(Sender: TObject);

    procedure bNextClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure bLastClick(Sender: TObject);
    procedure bPreviousClick(Sender: TObject);
    procedure bFirstClick(Sender: TObject);
    procedure bValWindowsClick(Sender: TObject);
    procedure bAnnulWindowsClick(Sender: TObject);
    procedure ToolbarButton971Click(Sender: TObject);
    procedure bValiderClick(Sender: TObject);
    procedure bAnnulerClick(Sender: TObject);
    Procedure CalculDuree(Sender: TObject);
    procedure Hdeb1Enter(Sender: TObject);
    procedure Hdeb1Exit(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure bDetailClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure HelpBtnClick(Sender: TObject);
  private

    StandardCal : string;    // Standard de calendrier associé
    FTypeHoraire:String;
    OldValue : double;
    Erreur : Boolean;
    Calendrier : TAFO_Calendrier;
    GrCalendrier : TGridCalendrierRessource;

  public
    { Déclarations publiques }
  end;

Procedure CalendrierDetail(TypeHoraire:String; Calendrier:TAFO_Calendrier;CodeStd,LibStd,CodeRes,CodeSalarie,LibNom:String);

implementation

{$R *.DFM}
Procedure CalendrierDetail(TypeHoraire:String; Calendrier:TAFO_Calendrier;CodeStd,LibStd,CodeRes,CodeSalarie,LibNom:String);
var FCalendrier : TFCalendrier;
begin
FCalendrier:=TFCalendrier.Create(Application);
FCalendrier.Calendrier:=Calendrier;
FCalendrier.FTypeHoraire:=TypeHoraire;
FCalendrier.StandardCal := CodeStd; FCalendrier.NomStandard.Caption := LibStd;
If (TypeHoraire ='RES') Then Begin FCalendrier.NomStandard.Caption:=CodeRes+' '+ LibNom; End
Else If (TypeHoraire ='SAL') Then Begin FCalendrier.NomStandard.Caption:=CodeSalarie+' '+ LibNom; End;
   Try
       FCalendrier.ShowModal;
   Finally
       FCalendrier.Free;
   End;
    Screen.Cursor :=SyncrDefault;
End;

procedure TFCalendrier.FormShow(Sender: TObject);
Var DefRect : TDefRect;
col,row : integer;
Ret : Boolean;
BEGIN
DefRect.NbLigneInRect := 3;
DefRect.NbColInRect := 1;
DefRect.IsBold := True;
DefRect.Cadrage := taCenter;
If(FTypeHoraire='RES') Then LabelStandard.Caption:=TraduitGA('Ressource');
GrCalendrier:=TGridCalendrierRessource.Create(Cal,Self,1,6,True,True,False,True,Now,LabelMois,DefRect);
COl := GrCalendrier.Grcal.Col; // mcd 16/06/03 force CellEnter sur jour positionner par défaut pour Ok des valeurs si modif 
Row := GrCalendrier.Grcal.Row;
GRCalendrier.CalendrierCellEnter( GrCalendrier,Col,row,Ret);
end;

procedure TFCalendrier.bValWindowsClick(Sender: TObject);
Var Horaire : TAFO_Horaire;
    Ferie : Boolean;
    FDeb1,FFin1,FDeb2,FFin2,FDuree : double;
begin
THCritMaskEdit(HDuree).SetFocus;
If(Not StrTimeIsNull(Hdeb1.Text))Then FDeb1:=StrTimeToFloat(Hdeb1.Text)else Fdeb1 :=0;
If(Not StrTimeIsNull(Hfin1.Text))Then FFin1:=StrTimeToFloat(Hfin1.Text)else FFin1 :=0;
If(Not StrTimeIsNull(Hdeb2.Text))Then FDeb2:=StrTimeToFloat(Hdeb2.Text)else Fdeb2 :=0;
If(Not StrTimeIsNull(Hfin2.Text))Then FFin2:=StrTimeToFloat(Hfin2.Text)else FFin2 :=0;
If(Not StrTimeIsNull(HDuree.Text))Then FDuree:=StrTimeToFloat(HDuree.Text)else FDuree :=0;
If (CheckTravailFerie.Checked=True) Then Ferie:=True else Ferie:=False;

Horaire:=TAFO_Horaire(GrCalendrier.GrCal.Objects[Cal.Col,Cal.Row]);
if Horaire = nil then Horaire:=TAFO_Horaire.Create;
Horaire.HDeb1:=FDeb1;  Horaire.Hfin1:=FFin1;
Horaire.HDeb2:=FDeb2;  Horaire.Hfin2:=FFin2;
Horaire.HDuree:=FDuree;
Horaire.TravailFerie:=Ferie;

GrCalendrier.GrCal.Objects[Cal.Col,Cal.Row]:=Horaire;
// Traitement de la TOB Détail
Calendrier.SetCalendrierDetail(Horaire,GrCalendrier.DateSelect(GrCalendrier.FDateDeb,Cal.Col,Cal.Row));
end;

procedure TFCalendrier.FormDestroy(Sender: TObject);
begin
GrCalendrier.Free
end;

// *** Gestion des boutons suivants , précédents ...
procedure TFCalendrier.bNextClick(Sender: TObject);
begin
If Calendrier.IsModifCalendrierDetail Then
    Begin
    If (PGIAsk('Voulez-vous enregistrer les modifications','Ressources')= mrYes) Then
        Calendrier.UpdateCalendrierDetail;
    End;

If (GrCalendrier.FCalendMois) Then
    Begin
    DecodeDate(PlusDate(EncodeDate(GrCalendrier.FYear,GrCalendrier.FMonth,1),1,'M'),GrCalendrier.FYear,GrCalendrier.FMonth,GrCalendrier.FDay);
    GrCalendrier.ChargeCalendrier
    End
Else
    Begin
    DecodeDate(PlusDate(GrCalendrier.DateSelect(GrCalendrier.FDateDeb,GrCalendrier.FColDebJour,Cal.TopRow),1,'M'),GrCalendrier.FYear,GrCalendrier.FMonth,GrCalendrier.FDay);
    Cal.TopRow:=((Trunc(EncodeDate(GrCalendrier.FYear,GrCalendrier.FMonth,GrCalendrier.FDay)-GrCalendrier.FDateDeb)) div 7)+1;
    End;
end;

procedure TFCalendrier.bLastClick(Sender: TObject);
begin
If Calendrier.IsModifCalendrierDetail Then
    Begin
    If (PGIAsk('Voulez-vous enregistrer les modifications','Ressources')= mrYes) Then
        Calendrier.UpdateCalendrierDetail;
    End;

If (GrCalendrier.FCalendMois) Then
    Begin
    DecodeDate(PlusDate(EncodeDate(GrCalendrier.FYear,GrCalendrier.FMonth,1),1,'A'),GrCalendrier.FYear,GrCalendrier.FMonth,GrCalendrier.FDay);
    GrCalendrier.ChargeCalendrier
    End
Else
    Begin
    DecodeDate(PlusDate(GrCalendrier.DateSelect(GrCalendrier.FDateDeb,GrCalendrier.FColDebJour,Cal.TopRow),1,'A'),GrCalendrier.FYear,GrCalendrier.FMonth,GrCalendrier.FDay);
    Cal.TopRow:=((Trunc(EncodeDate(GrCalendrier.FYear,GrCalendrier.FMonth,GrCalendrier.FDay)-GrCalendrier.FDateDeb)) div 7)+1;
    End;
end;

procedure TFCalendrier.bPreviousClick(Sender: TObject);
begin
If Calendrier.IsModifCalendrierDetail Then
    Begin
    If (PGIAsk('Voulez-vous enregistrer les modifications','Ressources')= mrYes) Then
        Calendrier.UpdateCalendrierDetail;
    End;

If (GrCalendrier.FCalendMois) Then
    Begin
    DecodeDate(PlusDate(EncodeDate(GrCalendrier.FYear,GrCalendrier.FMonth,1),-1,'M'),GrCalendrier.FYear,GrCalendrier.FMonth,GrCalendrier.FDay);
    GrCalendrier.ChargeCalendrier
    End
Else
    Begin
    DecodeDate(PlusDate(GrCalendrier.DateSelect(GrCalendrier.FDateDeb,GrCalendrier.FColDebJour,Cal.TopRow),-1,'M'),GrCalendrier.FYear,GrCalendrier.FMonth,GrCalendrier.FDay);
    Cal.TopRow:=((Trunc(EncodeDate(GrCalendrier.FYear,GrCalendrier.FMonth,GrCalendrier.FDay)-GrCalendrier.FDateDeb)) div 7)+1;
    End;
end;

procedure TFCalendrier.bFirstClick(Sender: TObject);

begin
If Calendrier.IsModifCalendrierDetail Then
    Begin
    If (PGIAsk('Voulez-vous enregistrer les modifications','Ressources')= mrYes) Then
        Calendrier.UpdateCalendrierDetail;
    End;

If (GrCalendrier.FCalendMois) Then
    Begin
    DecodeDate(PlusDate(EncodeDate(GrCalendrier.FYear,GrCalendrier.FMonth,1),-1,'A'),GrCalendrier.FYear,GrCalendrier.FMonth,GrCalendrier.FDay);
    GrCalendrier.ChargeCalendrier
    End
Else
    Begin
    DecodeDate(PlusDate(GrCalendrier.DateSelect(GrCalendrier.FDateDeb,GrCalendrier.FColDebJour,Cal.TopRow),-1,'A'),GrCalendrier.FYear,GrCalendrier.FMonth,GrCalendrier.FDay);
    Cal.TopRow:=((Trunc(EncodeDate(GrCalendrier.FYear,GrCalendrier.FMonth,GrCalendrier.FDay)-GrCalendrier.FDateDeb)) div 7)+1;
    End;
end;

procedure TFCalendrier.bAnnulWindowsClick(Sender: TObject);
begin
FenHoraire.hide;
end;

procedure TFCalendrier.ToolbarButton971Click(Sender: TObject);
begin
Close;
end;
procedure TFCalendrier.bValiderClick(Sender: TObject);
begin
Calendrier.UpdateCalendrierDetail;
end;

procedure TFCalendrier.bAnnulerClick(Sender: TObject);
begin
If Calendrier.IsModifCalendrierDetail Then
    Begin
    If (PGIAsk('Voulez-vous enregistrer les modifications','Ressources')= mrYes) Then
        Calendrier.UpdateCalendrierDetail;
    End;
Close;
end;

//////////// Partie Spécifique //////////////
Procedure TGridCalendrierRessource.ChargeSpecifGridCalendrier;
Begin
// Fonction de chargement des objets à associer au Grid sur la période
If (Ecran is TFCalendrier) then
    Begin
    TFCalendrier(Ecran).FenHoraire.hide;
    TFCalendrier(Ecran).Calendrier.ChargeCalendrierDetail(FDateDeb,FDateFin);
    End;
End;

Procedure TGridCalendrierRessource.ObjetCalendrierJour (LeJour : TDateTime;Col,Row,NumSemaine :integer);
Var Horaire : TAFO_Horaire;
Begin
// Affectation des objets à chaque Jour du Grid.
If (Ecran is TFCalendrier) then
    Begin
    Horaire:=TFCalendrier(Ecran).Calendrier.GetHoraire(LeJour);
    if (Horaire <> Nil) Then GrCal.Objects[Col,Row]:=Horaire;
    End;
End;

Procedure TGridCalendrierRessource.CalendrierDrawRect(Col,Row : integer;Rect:TRect;Ferie:Boolean);
Var Horaire : TAFO_Horaire;
    R : TRect;
    Horaire1,Horaire2: string;
    RazRect : Boolean;
Begin
// Dessin pour chaque cellule des zones complémentaires au numéro de jour (géré en standard)
R:= Rect; 
Horaire:=TAFO_Horaire(GrCal.Objects[Col,Row]);

RazRect := False;
If (Ferie) And ( (Horaire=nil) Or Not(Horaire.TravailFerie) ) then RazRect := True;
// Affichage horaires
if (Horaire<>nil) Or (RazRect)   then
begin
    if (Horaire<>nil)  And Not(RazRect) then
    BEGIN
    If ((Horaire.HDeb1 <> 0) and (Horaire.HFin1<>0)) Then
        Horaire1:= FloatToStrTime (Horaire.HDeb1, FormatHM) +' - '+ FloatToStrTime ( Horaire.HFin1, FormatHM) Else Horaire1:='';
END else Horaire1 := '';
// Ecritrue partie 2 du Rect
R.Top:=  Rect.Top+ ((Rect.Bottom - Rect.Top) div FDefRect.NbLigneInRect);
R.Bottom:=R.Top+((Rect.Bottom - Rect.Top) div 3);
    GrCal.canvas.TextRect(R, R.Left + (R.Right - R.Left - GrCal.canvas.TextWidth(Horaire1)) div 2,
    R.Top + (R.Bottom - R.Top - GrCal.canvas.TextHeight(Horaire1)) div 2, Horaire1);

if (Horaire<>nil) And Not(RazRect) then
BEGIN
    If ((Horaire.HDeb2 <> 0) and (Horaire.HFin2<>0)) Then
        Horaire2 := FloatToStrTime(Horaire.HDeb2 , FormatHM ) + ' - ' + FloatToStrTime (Horaire.HFin2, FormatHM) Else Horaire2:='';
END else Horaire2 := '';
// Ecriture Ligne 3 du Rect
R.Top:=R.Bottom;
R.Bottom:=R.Bottom+((Rect.Bottom - Rect.Top) div FDefRect.NbLigneInRect);
GrCal.canvas.TextRect(R, R.Left + (R.Right - R.Left - GrCal.canvas.TextWidth(Horaire2)) div 2,
    R.Top + (R.Bottom - R.Top - GrCal.canvas.TextHeight(Horaire2)) div 2, Horaire2);
end;

If (Ferie) And ( (Horaire=nil) Or Not(Horaire.TravailFerie) ) then
    // Dessin d'une croix sur les jours fériés
    Begin
    GrCal.Canvas.Pen.Color:= clRed;
    GrCal.canvas.MoveTo(Rect.left,Rect.top); GrCal.canvas.LineTo(Rect.right,Rect.bottom);
    GrCal.canvas.MoveTo(Rect.right,Rect.top); GrCal.canvas.LineTo(Rect.left,Rect.bottom);
    End;
End;


procedure TGridCalendrierRessource.CalendrierDblClick(Sender: TObject);
begin
If (Ecran is TFCalendrier) then
    begin
    TFCalendrier(Ecran).FenHoraire.Visible:=True ;
    THCritMaskEdit(TFCalendrier(Ecran).HDeb1).SetFocus;
    end;
end;

procedure TGridCalendrierRessource.CalendrierCellEnter(Sender: TObject; var ACol, ARow: Integer;  var Cancel: Boolean);
var      Horaire : TAFO_Horaire;
         IYear,IMonth,IDay : Word;
begin
If (Ecran is TFCalendrier) then
    BEGIN
    TFCalendrier(Ecran).Hdeb1.Text:=''; TFCalendrier(Ecran).Hfin1.Text := '';
    TFCalendrier(Ecran).Hdeb2.Text:=''; TFCalendrier(Ecran).Hfin2.Text := '';
    TFCalendrier(Ecran).HDuree.Text:='';TFCalendrier(Ecran).CheckTravailFerie.Checked:=False;

    DecodeDate(DateSelect(FDateDeb,GrCal.Col,GrCal.Row),IYear,IMonth,IDay);
    if TFCalendrier(Ecran).Calendrier.JourFerie.TestJourFerie(EncodeDate(IYear,IMonth,IDay)) then
        TFCalendrier(Ecran).CheckTravailFerie.Visible := true
    else TFCalendrier(Ecran).CheckTravailFerie.Visible := false;
    
    Horaire:=TAFO_Horaire(GrCal.Objects[GrCal.Col,GrCal.Row]);
    if IDay>0 then TFCalendrier(Ecran).FenHoraire.Caption:='Horaire du '+DateToStr(EncodeDate(IYear,IMonth,IDay))
              else TFCalendrier(Ecran).FenHoraire.Caption:='Horaire';
    if Horaire<>nil then
        BEGIN
        if (Horaire.HDeb1<> 0)   then TFCalendrier(Ecran).Hdeb1.Text := FloatToStrTime (Horaire.HDeb1, FormatHM );
        if (Horaire.Hfin1<> 0)   then TFCalendrier(Ecran).Hfin1.Text := FloatToStrTime (Horaire.Hfin1, FormatHM );
        if (Horaire.HDeb2<> 0)   then TFCalendrier(Ecran).Hdeb2.Text := FloatToStrTime (Horaire.HDeb2, FormatHM );
        if (Horaire.Hfin2<> 0)   then TFCalendrier(Ecran).Hfin2.Text := FloatToStrTime (Horaire.Hfin2, FormatHM );
        if (Horaire.HDuree<> 0)  then TFCalendrier(Ecran).Hduree.Text:= FloatToStrTime (Horaire.HDuree, FormatHM );

        if (Horaire.TravailFerie)then TFCalendrier(Ecran).CheckTravailFerie.Checked:=True
                                 else TFCalendrier(Ecran).CheckTravailFerie.Checked:=False;
        END;
    END;
END;

Procedure  TFCalendrier.CalculDuree(Sender: TObject);
Var //CDeb1, CFin1, CDeb2, CFin2 : THEDIT;
    Plage1, Plage2 : double;
BEGIN
Plage1 := 0 ;  Plage2 := 0;
if  (HDeb2.Text < HFin1.Text) And (HDeb2.Text <> '') And ( StrTimeToFloat(HDeb2.Text)<> 0.0) then
    BEGIN
    PgiBoxAF('L''horaire de début de la plage Horaire 2 est inférieur à la plage 1',TitreHalley);
    THCritMaskEdit(Sender).Text := ''; THCritMaskEdit(Sender).SetFocus;
    Erreur := true; Exit;
    END;

if  (HFin1.Text <> '') and (HDeb1.Text <> '') then
    Plage1 := CalculEcartHeure ( StrTimeToFloat (HDeb1.Text) , StrTimeToFloat (HFin1.Text) );
if  (HFin2.Text <> '') and (HDeb2.Text <> '') then
    Plage2 := CalculEcartHeure ( StrTimeToFloat (HDeb2.Text) , StrTimeToFloat (HFin2.Text) );

Plage1 := AjouteHeure (Plage1, Plage2);
HDuree.Text:= FloatToStrTime (Plage1 , FormatHM);
END;

procedure TFCalendrier.Hdeb1Enter(Sender: TObject);
begin
if (THEdit(Sender).Text <> '') then OldValue := StrTimeToFloat(THEdit(Sender).Text);
// Affichage au format double
THEdit(Sender).Text := FloatToStr( StrTimeToFloat (THEdit(Sender).Text) );
if Erreur then BEGIN OldValue := -1; Erreur := False; END;
end;

procedure TFCalendrier.Hdeb1Exit(Sender: TObject);
Var Tmp : Double;
begin
// mcd 04/10/02 car si Suppr sur le champ cela ne marche pas if (THEdit(Sender).Text ='') then Exit;
if (THEdit(Sender).Text ='') then THEdit(Sender).Text:='0';
Tmp := StrToFloat(THEdit(Sender).Text);
if Not TestHeure (tmp) then BEGIN THEdit(Sender).Text := ''; THEdit(Sender).SetFocus; Exit; END
                       else THEdit(Sender).Text := FloatToStrTime(tmp,FormatHM);
if (OldValue <> Tmp) then CalculDuree(Sender);
end;

procedure TFCalendrier.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
Var OkG,Vide : Boolean ;
begin
OkG:=(Screen.ActiveControl=Cal) ; Vide:=(Shift=[]) ;
Case Key of
   VK_RETURN : Key:=VK_TAB ;
   VK_F5     : if ((OkG) and (Vide)) then BEGIN Key:=0 ; GrCalendrier.CalendrierDblClick(Nil); END ;
   VK_INSERT : if ((OkG) and (Vide)) then BEGIN Key:=0 ; GrCalendrier.CalendrierDblClick(Nil); END ;
   VK_F10    : if Vide then BEGIN Key:=0 ; BValiderClick(Nil) ; END ;
  end ;
end;

procedure TFCalendrier.bDetailClick(Sender: TObject);
begin
GrCalendrier.CalendrierDblClick(Nil);
end;


procedure TFCalendrier.BImprimerClick(Sender: TObject);
begin
LanceEtat ('E','ACE','ACE',True,False,False,Nil,'','',False);
end;

procedure TFCalendrier.HelpBtnClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.
