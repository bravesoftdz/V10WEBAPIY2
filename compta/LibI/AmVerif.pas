
unit AmVerif;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Vierge, ExtCtrls, HPanel, HSysMenu, HTB97, StdCtrls, ColMemo, ImgList,
  HMsgBox, hctrls, utob
  {$IFDEF EAGLCLIENT}
  {$ELSE}
  {$IFNDEF DBXPRESS} ,dbtables {$ELSE} ,uDbxDataSet {$ENDIF}
  {$ENDIF}
  , AmRapport
  ;


type
  TFVerifAmort = class(TFVierge)
    HPanel1: THPanel;
    LesImages: TImageList;
    FCBVerif1: TCheckBox;
    FBCheck1: TToolbarButton97;
    FBDetail1: TToolbarButton97;
    FMemoDetail: TColorMemo;
    FCBVerif2: TCheckBox;
    FBCheck2: TToolbarButton97;
    FBDetail2: TToolbarButton97;
    FCBVerif4: TCheckBox;
    FBCheck4: TToolbarButton97;
    FBDetail4: TToolbarButton97;
    FCBVerif5: TCheckBox;
    FBCheck5: TToolbarButton97;
    FBDetail5: TToolbarButton97;
    FCBVerif6: TCheckBox;
    FBCheck6: TToolbarButton97;
    FBDetail6: TToolbarButton97;
    FCBVerif7: TCheckBox;
    FBCheck7: TToolbarButton97;
    FBDetail7: TToolbarButton97;
    BToutSel: TToolbarButton97;
    FBCheck3: TToolbarButton97;
    FBDetail3: TToolbarButton97;
    FCBVerif3: TCheckBox;
    FCBVerif8: TCheckBox;
    FBCheck8: TToolbarButton97;
    FBDetail8: TToolbarButton97;
    FCBVerif9: TCheckBox;
    FBCheck9: TToolbarButton97;
    FBDetail9: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure FCBVerif1MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure HPanel1MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure BValiderClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FBDetailClick(Sender: TObject);
    procedure BToutSelClick(Sender: TObject);
    procedure FCBVerif4Click(Sender: TObject);
  private
    { Déclarations privées }
    fListeVerifAmort  :  array of TActionVerifAmort;
    fRapportVerif : TOB;
    procedure InitLesBoutons;
  public
    { Déclarations publiques }
  end;

// MVG 12/07/2006
{$IFDEF SERIE1}
procedure LanceVerificationAmortissement(Inside:ThPanel);
{$ELSE}
procedure LanceVerificationAmortissement;
{$ENDIF SERIE1}

implementation

{$IFDEF SERIE1}
uses hent1, UIUtil ;
{$ELSE}
{$ENDIF SERIE1}
{$IFDEF A_SUPPRIMER?}
uses ImEnt;
{$ENDIF !A_SUPPRIMER?}

{$R *.DFM}

{$IFDEF SERIE1}
// MVG 12/07/2006
procedure LanceVerificationAmortissement(Inside:ThPanel);
var x:TFVerifAmort;
begin
   SourisSablier;
   x:=TFVerifAmort.Create(Application) ;
   If Inside=nil then
   begin
      Try
         x.showmodal;
       finally
         x.Free;
      end;
   end else
   begin
      InitInside (X,Inside);
      x.Show;
   end;
   SourisNormale;
end;
{$ELSE}
procedure LanceVerificationAmortissement;
var
  FVerifAmort: TFVerifAmort;
begin
  FVerifAmort := TFVerifAmort.Create(Application);
  try
    FVerifAmort.ShowModal;
  finally
    FVerifAmort.Free;
  end;
end;
{$ENDIF}

procedure TFVerifAmort.InitLesBoutons;
var i : integer;
    Bouton : TToolbarButton97;
begin
  for i:= 1 to MAX_ACTION do
  begin
    Bouton := TToolbarButton97(FindComponent('FBCheck' + IntToStr(i)));
    if Bouton <> nil then Bouton.Visible := False;
    Bouton := TToolbarButton97(FindComponent('FBDetail' + IntToStr(i)));
    if Bouton <> nil then Bouton.Visible := False;
  end;
end;

procedure TFVerifAmort.FormShow(Sender: TObject);
begin
  inherited;
  FMemoDetail.Lines.Clear;
  InitLesBoutons;
  SetLength(fListeVerifAmort, MAX_ACTION+1);
  ChargeListeActionVerifAmort (fListeVerifAmort) ;
end;

procedure TFVerifAmort.FCBVerif1MouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
var i : integer;
    St : string;
begin
  inherited;
  St := TToolbarButton97(Sender).Name;
  i := StrToInt(Copy(St,9,Length(St)-8));
  if FMemoDetail.Lines.Count=0 then FMemoDetail.Lines.Add(fListeVerifAmort[i].Info);
end;

procedure TFVerifAmort.HPanel1MouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  inherited;
  FMemoDetail.Lines.Clear;
end;

procedure TFVerifAmort.BValiderClick(Sender: TObject);
var Action : TCheckBox;
    Bouton : TToolbarButton97;
    i : integer;
    bOkVerif : boolean;
    TRapport : TOB;
begin
  inherited;
  InitLesBoutons;
  fRapportVerif.ClearDetail;
  for i := 1 to MAX_ACTION do
  begin
    Action := TCheckBox(FindComponent('FCBVerif' + IntToStr(i)));
    if Action <> nil then
    begin
      TRapport := TOB.Create('',fRapportVerif,-1);
      if Action.Checked then
      begin
        fListeVerifAmort[i].ToDo( bOkVerif , TRapport , True);
        Bouton := TToolbarButton97(FindComponent('FBCheck' + IntToStr(i)));
        if Bouton <> nil then
        begin
          Bouton.Visible := True;
          if bOkVerif then Bouton.ImageIndex := 0
          else Bouton.ImageIndex := 1;
          if fRapportVerif.Detail[i-1].Detail.Count > 0 then
          begin
            Bouton := TToolbarButton97(FindComponent('FBDetail' + IntToStr(i)));
            if Bouton <> nil then Bouton.Visible := True;
          end;
        end;
      end;
    end;
  end;
end;

procedure TFVerifAmort.FormCreate(Sender: TObject);
begin
  inherited;
  fRapportVerif := TOB.Create ('', nil, -1);
  {$IFDEF SERIE1}
  HelpContext:=537000 ; //YCP 17/10/2006 helpcontext
  {$ELSE}
  {$ENDIF SERIE1}
end;

procedure TFVerifAmort.FormDestroy(Sender: TObject);
begin
  fRapportVerif.Free;
  inherited;
end;

procedure TFVerifAmort.FBDetailClick(Sender: TObject);
var i : integer;
    St : string;
    stCaption : string;
begin
  inherited;
  St := TToolbarButton97(Sender).Name;
  i := StrToInt(Copy(St,9,Length(St)-8));
  stCaption := TCheckBox(FindComponent('FCBVerif'+IntToStr(i))).Caption;
  AfficheRapportVerifAmort ( i, fRapportVerif.Detail[i-1] , fListeVerifAmort, stCaption);
end;

procedure TFVerifAmort.BToutSelClick(Sender: TObject);
var i : integer;
    Action : TCheckBox;
    bSelect : boolean;
begin
  inherited;
  bSelect :=  FCBVerif1.Checked;
  for i := 1 to MAX_ACTION do
  begin
    Action := TCheckBox(FindComponent('FCBVerif' + IntToStr(i)));
    if Action <> nil then Action.Checked := not bSelect;
  end;
end;

procedure TFVerifAmort.FCBVerif4Click(Sender: TObject);
begin
  inherited;
  if FCBVerif4.Checked then
  begin
    FCBVerif3.Enabled := False;
    FCBVerif3.Checked := True;
  end else FCBVerif3.Enabled := True; 
end;

end.
