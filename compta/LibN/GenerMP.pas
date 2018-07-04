unit GenerMP;

interface

uses
  SysUtils, WinTypes, WinProcs, Messages, Classes, Graphics, Controls,HEnt1,
  Forms, Dialogs, StdCtrls, Mask, Hctrls, Buttons, ExtCtrls, hmsgbox,Ent1, DBTables,
  GuidTool, {FP 31/05/2006 FQ17982}
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  Saisutil, LettUtil, HSysMenu ;


Type tGenereMP = Record
                 CptG,MPG,JalG,NatureG,CptS,NatJalG : String ;
                 GroupeEncadeca,ModeAnal,NumEncaDeca : String ;
                 BanquePrevi : string; {JP 22/02/05 : Gestion des banques prévisionnelles}
                 DateC,DateE : tDateTime ;
                 ForceEche,GenColl,GenPoint,GenLett,GenVent : Boolean ;
                 Ventils : Array[1..MaxAxe] of boolean ;
                 Ref1,Lib1,Ref2,Lib2,RefExt1,RefLib1,RefExt2,RefLib2 : String ;
                 Cat,NomFSelect : String ;
                 DEV            : RDEVISE ;
                 smp            : TSuiviMP ;
                 ExportCFONB,EnvoiBordereau : boolean ;
                 FormatCFONB,ModeleBordereau : String ;
                 ModeTeletrans,ChampsRupt : String ;
                 AlerteEcheMP : Boolean ;
                 TIDTIC : Boolean ;
                 LettrAuto : Boolean ; // YMO 25/11/2005 : Lettrage désactivé pour reglt en devises
                 VisuEcrGene: Boolean;  {FP 25/04/2006 FQ17966: Permet de voir les écritures générées lors de la compensation}
                 CumulEcrGene: Boolean;  {FP 01/06/2006 FQ18193: Permet de cumuler les écritures à générer}
                 TauxVolatil : Boolean; {JP 13/06/06 : FQ 18319 : Si un taux volatil a été saisi dans GenereMP}
                 End ;

Function  ParamsMPSup(Var GMP : tGenereMP ; SetFocusMP : Boolean = FALSE) : Boolean ;

type
  TFGenereMPPlus = class(TForm)
    BValider: THBitBtn;
    BAnnuler: THBitBtn;
    Baide: THBitBtn;
    HM: THMsgBox;
    HMTrad: THSystemMenu;
    HDateReg: THLabel;
    DATEGENERATION: THCritMaskEdit;
    CChoixDate: TCheckBox;
    DATEECHEANCE: THCritMaskEdit;
    HLabel4: THLabel;
    NUMENCADECA: THCritMaskEdit;
    H_MODEGENERE: TLabel;
    MODEGENERE: THValComboBox;
    CCVisuEcr: TCheckBox;
    CCCumulEcr: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BValiderClick(Sender: TObject);
    procedure BaideClick(Sender: TObject);
    procedure CChoixDateClick(Sender: TObject);
    procedure MODEGENEREChange(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure NUMENCADECAElipsisClick(Sender: TObject);     {FP 31/05/2006 FQ17982}
  private
    GMP : tGenereMP ;
    SetFocusMP : Boolean ;
    Function DateOk : Boolean ;
  public
  end;

implementation

{$R *.DFM}

Uses
  EncUtil, CPCHANCELL_TOF {FQ 16617};

{JP 13/06/06 : FQ 18319 : j'ai bien du mal à comprendre pourquoi on s'embête à passer certaines
               valeurs du GMP et non pas toute la structure => risque de perte d'informations.
               Je prends le parti de simplifier et de réduire cela à deux lignes de code
{---------------------------------------------------------------------------------------}
Function  ParamsMPSup(Var GMP : tGenereMP ; SetFocusMP : Boolean = FALSE) : Boolean ;
{---------------------------------------------------------------------------------------}
var
  FGenereMPPlus : TFGenereMPPlus ;
begin
  FGenereMPPlus := TFGenereMPPlus.Create(Application);
  try
    FGenereMPPlus.GMP := GMP; {JP 13/06/06 : FQ 18319}
    (*
    FGenereMPPlus.GMP.VisuEcrGene:=GMP.VisuEcrGene;{FP 25/04/2006 FQ17966}
    FGenereMPPlus.GMP.CumulEcrGene:=GMP.CumulEcrGene; {FP 01/06/2006 FQ18193}
    FGenereMPPlus.GMP.DateC:=GMP.DateC ;
    FGenereMPPlus.GMP.DateE:=GMP.DateE ;
    FGenereMPPlus.GMP.NumEncadeca:=GMP.NumEncadeca ;
    FGenereMPPlus.GMP.ForceEche:=GMP.ForceEche ;
    FGenereMPPlus.GMP.MPG:=GMP.MPG ;
    FGenereMPPlus.GMP.Cat:=GMP.Cat ;
    FGenereMPPlus.GMP.AlerteEcheMP:=GMP.AlerteEcheMP ;
    FGenereMPPlus.SetFocusMP:=SetFocusMP ;
    FGenereMPPlus.GMP.smp:= GMP.smp;              {FP 21/02/2006}

    {JP 28/09/05 : FQ 16617 : Initialisation de la structure devise}
    FGenereMPPlus.GMP.DEV.Code := GMP.Dev.Code;
    FGenereMPPlus.GMP.DEV.Taux := 1;
    FGenereMPPlus.GMP.DEV.DateTaux := iDate1900;
    *)
    Result:=(FGenereMPPlus.ShowModal=mrOK) ;
    if Result then
      GMP := FGenereMPPlus.GMP; {JP 13/06/06 : FQ 18319}
      (*
      GMP.VisuEcrGene:=FGenereMPPlus.GMP.VisuEcrGene;{FP 25/04/2006 FQ17966}
      GMP.CumulEcrGene:=FGenereMPPlus.GMP.CumulEcrGene; {FP 01/06/2006 FQ18193}
      GMP.DateC:=FGenereMPPlus.GMP.DateC ;
      GMP.DateE:=FGenereMPPlus.GMP.DateE ;
      GMP.NumEncadeca:=FGenereMPPlus.GMP.NumEncadeca ;
      GMP.ForceEche:=FGenereMPPlus.GMP.ForceEche ;
      GMP.MPG:=FGenereMPPlus.GMP.MPG ;

      {JP 28/09/05 : FQ 16617 : éventuelle récupération des données}
      if (FGenereMPPlus.GMP.DEV.DateTaux > iDate1900) and
         (GMP.Dev.Code <> v_PGI.DevisePivot) and
         (FGenereMPPlus.GMP.DEV.Taux <> 1) then begin
        GMP.DEV.Taux     := FGenereMPPlus.GMP.DEV.Taux;
        GMP.DEV.DateTaux := FGenereMPPlus.GMP.DEV.DateTaux;
      end;
      *)
  finally
    FGenereMPPlus.Free ;
  end;
  Screen.Cursor:=SyncrDefault ;
END ;

procedure TFGenereMPPlus.FormShow(Sender: TObject);
Var i : Integer ;
    OffsetWidth:  Integer;   {FP 25/04/2006 FQ17982}
    OffsetHeight: Integer;   {FP 01/01/2006 FQ18193}
    List:         TStrings;  {FP 01/01/2006 FQ18193}
begin
CatToMP(GMP.Cat,MODEGENERE.Items,MODEGENERE.Values,tslAucun,TRUE) ; ModeGenere.itemIndex:=0 ;
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
BValider.Visible:=TRUE ; BAnnuler.Visible:=TRUE ; BAide.Visible:=TRUE ;
BValider.Enabled:=TRUE ; BAnnuler.Enabled:=TRUE ; BAide.Enabled:=TRUE ;
NumEncaDeca.Text:=GMP.NumEncadeca ;
(*
If GMP.DateC=0 Then DateGeneration.Text:=DateToStr(V_PGI.DateEntree)
               Else DateGeneration.Text:=DateToStr(GMP.DateC) ;
*)
{b FP 21/02/2006}
if GMP.smp in [smpCompenCli, smpCompenFou] then
  begin
  {b FP 25/04/2006 FQ17982: Utilisation du composant NumEncaDeca pour stocker le libellé}
  OffsetWidth  := 25;
  OffsetHeight := 20;                                         {FP 01/01/2006 FQ18193}
  HLabel4.Caption  := TraduireMemoire('&Libellé');
  {b FP 31/05/2006 FQ17982}
  NUMENCADECA.DataType      := '';
  NUMENCADECA.ElipsisButton := True;
  NUMENCADECA.OnElipsisClick:= NUMENCADECAElipsisClick;
  {e FP 31/05/2006 FQ17982}
  NUMENCADECA.CharCase      := ecNormal;
  NUMENCADECA.MaxLength     := 35;
  CCVisuEcr.Visible         := True;
  CCCumulEcr.Visible        := True;
  {Agrandit la taille de la fenêtre}
  List := TStringList.Create;
  try                
    for i:=0 to ControlCount-1 do
      List.AddObject(IntToStr(Controls[i].Top), Controls[i]);
    Self.Width           := Self.Width           + 30;
    Self.Height          := Self.Height          + 25; {FP 01/01/2006 FQ18193}
    for i:=0 to List.Count-1 do
      (List.Objects[i] as TControl).Top := StrToInt(List[i]);
  finally
    List.Free;
    end;
  DATEGENERATION.Width := DATEGENERATION.Width + OffsetWidth;
  DATEECHEANCE.Width   := DATEECHEANCE.Width   + OffsetWidth;
  NUMENCADECA.Width    := NUMENCADECA.Width    + OffsetWidth;
  MODEGENERE.Width     := MODEGENERE.Width     + OffsetWidth;
  BValider.Left        := BValider.Left        + OffsetWidth;
  BValider.Top         := BValider.Top         + OffsetHeight; {FP 01/01/2006 FQ18193}
  BAnnuler.Left        := BAnnuler.Left        + OffsetWidth;
  BAnnuler.Top         := BAnnuler.Top         + OffsetHeight; {FP 01/01/2006 FQ18193}
  Baide.Left           := Baide.Left           + OffsetWidth;
  Baide.Top            := Baide.Top            + OffsetHeight; {FP 01/01/2006 FQ18193}
  {e FP 25/04/2006}
  end;
{e FP 21/02/2006}
CCVisuEcr.Checked  := GMP.VisuEcrGene;     {FP 25/04/2006 FQ17966}
CCCumulEcr.Checked := GMP.CumulEcrGene;    {FP 01/06/2006 FQ18193}
DateGeneration.Text:=DateToStr(V_PGI.DateEntree) ;
DateEcheance.Text:=DateToStr(GMP.DateE) ;
CChoixDate.Checked:=GMP.ForceEche ; CChoixDateClick(Nil) ;
If (GMP.MPG<>'') And (MODEGENERE.Values.Count>0) Then
  BEGIN
  i:=MODEGENERE.Values.IndexOf(GMP.MPG) ;
  If i>=0 Then MODEGENERE.ItemIndex:=i ;
  END ;
If SetFocusMP Then ModeGenere.SetFocus ;
end;

procedure TFGenereMPPlus.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
if Shift<>[] then Exit ;
if Key=VK_F10 then BEGIN Key:=0 ; BValiderClick(Nil) ; END else
 if Key=VK_ESCAPE then BEGIN Key:=0 ; ModalResult:=mrCancel ; END ;
end;

procedure TFGenereMPPlus.BValiderClick(Sender: TObject);
var
  Dt : TDatetime;
  Ds : TDatetime;
  Tx : Double;
begin
  If Not DateOk Then Exit ;

  Ds := StrToDate(DateGeneration.Text); {JP 13/10/05}
  {JP 13/06/06 : FQ 18319 : Pour ne pas reprendre le taux de la table Chancell dans GenerMP}
  if not GMP.TauxVolatil then begin
    {JP 28/09/05 : FQ 16617 : Gestion du taux de devise si changement de date comptable}
    if {(GMP.DateC <> Ds) and} (GMP.Dev.Code <> V_PGI.DevisePivot) then begin
      {JP 13/10/05if HSHowMessage('0;' + Caption + ';Voulez-vous mettre à jour le taux de la devise ?;Q;YN;Y;N;', '', '') = mrYes then begin}
        {On regarde si le taux de la devise existe pour la date comptable ...}
        Tx := GetTaux(GMP.Dev.Code, Dt, Ds{JP 13/10/05 : GMP.DateC});
        {... s'il n'y a pas de chancellerie à la date comptable ...}
        if (Dt <> Ds{JP 13/10/05 : GMP.DateC}) then begin
          {... On propose de saisir la chancellerie}
          if HSHowMessage('0;' + Caption + ';Voulez-vous saisir le taux de la devise ?;Q;YN;Y;N;', '', '') = mrYes then begin
            FicheChancel(GMP.Dev.Code, True, Ds{JP 13/10/05 : GMP.DateC}, taCreat, True);
            {On récupère le taux saisi}
            Tx := GetTaux(GMP.Dev.Code, Dt, Ds{JP 13/10/05 : GMP.DateC});
          end;
        end;
        {Si un taux a été récupéré, on met à jour la structure}
        if Tx <> 1 then begin
          GMP.Dev.Taux := Tx;
          GMP.Dev.DateTaux := Dt;
        end;
      {JP 13/10/05 end;}
    end;
  end;

  GMP.VisuEcrGene := CCVisuEcr.Checked;    {FP 25/04/2006 FQ17966}
  GMP.CumulEcrGene:= CCCumulEcr.Checked;   {FP 01/06/2006 FQ18193}
  GMP.DateC := StrToDate(DateGeneration.Text) ;
  GMP.DateE := StrToDate(DateEcheance.Text) ;
  GMP.NumEncadeca := NumEncaDeca.Text ;
  GMP.MPG := '' ;
  If ModeGenere.itemIndex > 0 Then GMP.MPG := MODEGENERE.Value ;
  GMP.ForceEche := CChoixDate.Checked ;
  ModalResult := mrOk ;
end;

procedure TFGenereMPPlus.BaideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

Function TFGenereMPPlus.DateOk : Boolean ;
Var DD  : TDateTime ;
    Err : integer ;
BEGIN
Result:=FALSE ;
if Not IsValidDate(DATEGENERATION.Text) then
   BEGIN
   HM.Execute(3,caption,'') ; DATEGENERATION.Text:=DateToStr(V_PGI.DateEntree) ; Exit ;
   END else
   BEGIN
   // FQ 17165 : contrôle fourchette de dates
   Err := ControleDate(DATEGENERATION.Text) ;
   if Err>0 then
      BEGIN
      HM.Execute(7+Err,caption,'') ;
      DATEGENERATION.Text:=DateToStr(V_PGI.DateEntree) ;
      Exit ;
      END else
      BEGIN
      DD  := StrToDate(DATEGENERATION.Text) ;
      if RevisionActive(DD) then
         BEGIN
         DATEGENERATION.Text:=DateToStr(V_PGI.DateEntree) ;
         Exit ;
         END ;
      // FQ 17165 : contrôle fourchette de dates
      if CChoixDate.Checked then
        if not NbJoursOk( DD, StrToDate(DATEECHEANCE.Text) ) then
          begin
          HM.Execute(2,caption,'') ;
          DATEECHEANCE.SetFocus ;
          Exit;
          end ;
      END ;
   END ;
Result:=TRUE ;
END ;


procedure TFGenereMPPlus.CChoixDateClick(Sender: TObject);
begin
if CChoixDate.Checked then
   BEGIN
   DateEcheance.Enabled:=True ;
   if DateEcheance.Text=StDate1900 then DateEcheance.Text:=DateToStr(V_PGI.DateEntree) ;
   If ModeGenere.ItemIndex>0 Then GMP.AlerteEcheMP:=FALSE ;
   END else
   BEGIN
   DateEcheance.Text:=StDate1900 ;
   DateEcheance.Enabled:=False ;
   END ;
end;

procedure TFGenereMPPlus.MODEGENEREChange(Sender: TObject);
begin
If (MODEGENERE.ItemIndex>0) And CCHoixDate.Checked Then
  BEGIN
  GMP.AlerteEcheMP:=FALSE ;
  END ;
end;

procedure TFGenereMPPlus.BAnnulerClick(Sender: TObject);
begin
ModalResult:=mrCancel
end;

{b FP 31/05/2006 FQ17982}
procedure TFGenereMPPlus.NUMENCADECAElipsisClick(Sender: TObject);
var
  St, StC: string;
  TE: tEdit;
begin
  inherited;
  TE := TEdit(Sender);
  St := ChoixChampZone(0, 'LIB');
  if St = '' then Exit;
  Stc := TE.Text;
  if Length(Stc + ' ' + St) > 100 then HM.Execute(13, Caption, '');
  if TE.SelLength > 0 then Delete(StC, TE.SelStart + 1, TE.SelLength);
  if StC = '' then StC := St else StC := StC + ' ' + St;
  TE.Text := StC;
end;
{e FP 31/05/2006 FQ17982}
end.
