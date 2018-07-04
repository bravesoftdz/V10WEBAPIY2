{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 04/03/2003
Modifié le ... :   /  /    
Description .. : Passage en eAGL
Mots clefs ... : 
*****************************************************************}
unit CodeSect;

interface

uses
  Windows,
  Classes,
  Controls,
  Messages,  // Tmessage, WM_GetMinMaxInfo
  SysUtils,  // Trim
  Forms,     // TForm
  Grids,
  StdCtrls,
  ExtCtrls,  // TPanel
  Buttons,   // TBitBtn
  HEnt1,     // TabByte, TabSt3, TabSt35
  HMsgBox,   // THMsgBox
  Hctrls,    // THLabel
  HSysMenu,  // THSystemMenu
  Ent1,      // ChargeTablo, VH
  UTob       // Tob
  ;

Procedure FicheCodeSect(Axe : String ; Var Cod,Lib : String) ;

type
  TFCodesect = class(TForm)
    PBouton : TPanel;
    MsgBox : THMsgBox;
    Panel1 : TPanel;
    TLibSec : THLabel;
    LibSec : TEdit;
    ECodSec : TEdit;
    TCodSec : TLabel;
    Panel2 : TPanel;
    BAide: THBitBtn;
    BFerme: THBitBtn;
    BValider: THBitBtn;
    HMTrad : THSystemMenu;
    BSaiCodsec: THBitBtn;
    FListe : THGrid;
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);

    procedure ECodSecKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure ECodSecMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);

    procedure FListeDblClick(Sender: TObject);
    procedure FListeKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);

    procedure BValiderClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure BSaiCodsecClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
    UnAxe,Cod,Lib,MemoLib : String ;
    OuSuisJe : Byte ;
    LaLongueur : Byte ;
    EstStructurer : Boolean ;
    Lindice : Byte ;
    DebStruct : TabByte ;
    LonStruct : TabByte ;
    CodStruct : TabSt3 ;
    LibelleStruct : TabSt35 ;
    WMinX,WMinY : Integer ;
    TobQSSSTRUCR : Tob;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure OnChangeLaGrid ;
    Procedure RafraichiDBGrid(i : Byte) ;
    Function AuneSousSection( i : Byte):Byte ;
    Procedure EnvoiRafraichiGrid(j : Byte) ;
    Procedure TrouveNouveau ;
    Procedure Libelle(St : String) ;
  public
  { Déclarations publiques }
end;

implementation

{$R *.DFM}

Uses
{$IFDEF MODENT1}
  CPProcMetier,
{$ENDIF MODENT1}
{$IFDEF EAGLCLIENT}
  CPSTRUCTURE_TOF, // ParamPlanAnal
{$ELSE}
    // FQ 20198 Structur,      // ParamPlanAnal
{$IFNDEF CMPGIS35}
    CPSTRUCTURE_TOF,
{$ENDIF}
{$ENDIF}
  UtilPGI  // _Blocage
  ;

procedure TFCodesect.FormShow(Sender: TObject);
begin
  PopUpMenu := ADDMenuPop(PopUpMenu,'','') ;
  ChargeTablo(UnAxe,DebStruct,LonStruct,CodStruct,LibelleStruct) ;

  FListe.Cells[0,0] := 'Code';
  FListe.Cells[1,0] := 'Libellé';
  FListe.ColWidths[0] := 50;
  FListe.ColWidths[1] := 240;

  ECodSec.MaxLength := VH^.Cpta[AxeToFb(UnAxe)].Lg ;
  OnChangeLaGrid ;
end;

// Rafraichit l'affichage de la grille quand on se positionne à un endroit dans le contrôle "CODE"
// Se positionne ensuite dans la grille sur la sélection
Procedure TFCodesect.RafraichiDBGrid(i : Byte) ;
var
  St : String ;
  Trouver : Boolean ;
begin
  if (FListe.RowCount >= 2) then
    FListe.VidePile(False);
  TobQSSSTRUCR := Tob.Create('SSSTRUCR',nil,-1);
  TobQSSSTRUCR.LoadDetailDB('SSSTRUCR','"'+UnAxe+'";"'+CodStruct[i]+'"','',nil,False);
  TobQSSSTRUCR.PutGridDetail(FListe,False,False,'PS_CODE;PS_LIBELLE');
  TobQSSSTRUCR.Free;
  FListe.Refresh;

  if ECodSec.Text<>'' then begin
    St := Copy(ECodSec.Text,DebStruct[i],LonStruct[i]) ;
    Trouver := False ;

    for i := 1 to FListe.RowCount-1 do begin
      if FListe.Cells[0,i] = St then begin
        FListe.Row := i;
        Trouver := True;
        Break;
      end;
    end;
    if Not Trouver then FListe.Row := 1;
  end;
end;

Function TFCodesect.AuneSousSection( i : Byte):Byte ;
var
  j : Byte ;
  Trouver : Boolean ;
begin
  Trouver := False ;
  j := 1 ;
  While (j<=ECodsec.MaxLength) And (Not Trouver) do begin
    if (i in [DebStruct[j]..(DebStruct[j]+LonStruct[j])-1]) AND (((DebStruct[j]+LonStruct[j])-1)>0) then begin
      Trouver := True ;
      EstStructurer := True ;
      end
    else
      Inc(j) ;
  end;
  if Trouver then Result := j
             else begin EstStructurer := False ; Result := 0 ; end;
end;

Procedure TFCodesect.OnChangeLaGrid ;
var
  i : Byte ;
begin
  OuSuisJe := ECodSec.SelStart+1 ;
  i := AuneSousSection(OuSuisJe) ;
  if i>0 then begin
    Lindice := i ;
    TrouveNouveau ;
  end;
end;

Procedure TFCodesect.EnvoiRafraichiGrid(j : Byte) ;
begin
  RafraichiDBGrid(j) ;
  LaLongueur := LonStruct[j] ;
end;

Procedure TFCodesect.TrouveNouveau ;
var
  Dedans : Boolean ;
begin
  Dedans := False ;
  if Lindice>=1 then begin
    if OuSuisje in [DebStruct[Lindice]..(DebStruct[Lindice]+LonStruct[Lindice])-1] then begin
      EnvoiRafraichiGrid(Lindice) ;
      Dedans := True ;
    end;
    if Lindice>1 then begin
      if OuSuisje in [DebStruct[Lindice-1]..(DebStruct[Lindice-1]+LonStruct[Lindice-1])-1] then begin
        Dec(Lindice) ;
        EnvoiRafraichiGrid(Lindice) ;
        Dedans := True ;
      end
    end;
  end;
  if not Dedans then FListe.VidePile(True);
end;

procedure TFCodesect.ECodSecKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  Case Key of
    VK_HOME,
    VK_END,
    VK_RIGHT,
    VK_LEFT,
    VK_UP,
    VK_DOWN  : begin
               OuSuisJe := ECodSec.SelStart+1 ;
               Lindice := AuneSousSection(OuSuisJe) ;
               TrouveNouveau ;
    end;
  end;
end;

procedure TFCodesect.FListeKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key=VK_RETURN then FlisteDblClick(Nil) ;
end;

procedure TFCodesect.FListeDblClick(Sender: TObject);
var
  S : String ;
begin
  S := Copy(ECodSec.Text,1,(DebStruct[Lindice]-1)) ;
  S := S + FListe.Cells[0,FListe.Row];
  Libelle(FListe.Cells[1,FListe.Row]);

  ECodSec.Text := S+Copy(ECodSec.Text,Length(S)+1,Length(EcodSec.Text)-Length(S)) ;
  if Length(EcodSec.Text)=ECodSec.MaxLength then begin
    BValider.SetFocus ;
  end;
  ECodSec.Setfocus ;
  ECodSec.SelStart := DebStruct[Lindice]+LaLongueur+1 ;
  ECodSec.SelLength := 0 ;
  OuSuisje := ECodSec.SelStart+1 ;
  Inc(Lindice) ;
  TrouveNouveau ;
end ;

Procedure TFCodesect.Libelle(St : String) ;
var
  Stemp,StMemo,S1,S2 : String ;
  i,j : Byte ;
begin
  St := Trim(St) ;
  if LibSec.Text='' then
    MemoLib := St+';'
  else begin
    Stemp := MemoLib ;
    if Lindice=1 then begin
      Delete(Stemp,1,Pos(';',Stemp)) ;
      MemoLib := St+';'+Stemp ;
      end
    else begin
      StMemo := '' ;
      j := 0 ;
      for i := 1 to Length(Stemp) do begin
        if Stemp[i]=';' then Inc(j) ;
        if j=Lindice-1 then StMemo := StMemo+Stemp[i] ;
        if j=Lindice then Break ;
      end;
      Delete(StMemo,1,1) ;
      if StMemo='' then
        MemoLib := Stemp+St+';'
      else begin
        j := Pos(StMemo,Stemp) ;
        Delete(Stemp,j,Length(StMemo)) ;
        S1 := Copy(Stemp,1,j-1) ;
        S2 := Copy(Stemp,j,Length(Stemp)) ;
        Stemp := S1+St+S2 ;
        MemoLib := Stemp ;
      end;
    end;
  end;
  S1 := MemoLib ;
  While Pos(';',S1) > 0 do
    S1[Pos(';',S1)] := ' ';
  LibSec.Text := S1 ;
end;

procedure TFCodesect.ECodSecMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  OuSuisJe := ECodSec.SelStart+1 ;
  Lindice := AuneSousSection(OuSuisJe) ;
  TrouveNouveau ;
end;

procedure TFCodesect.BFermeClick(Sender: TObject);
begin
  Cod := '' ;
  Lib := '' ;
  Close;
end;

procedure TFCodesect.BValiderClick(Sender: TObject);
begin
  Cod := ECodSec.Text ;
  Lib := MemoLib ;
  if Length(Cod)>ECodSec.MaxLength then Cod := Copy(Cod,1,ECodSec.MaxLength) ;
  Close ;
end;

procedure TFCodesect.BSaiCodsecClick(Sender: TObject);
begin
{$IFNDEF CMPGIS35}
  ParamPlanAnal(UnAxe) ;
  ChargeTablo(UnAxe,DebStruct,LonStruct,CodStruct,LibelleStruct) ;
  OnChangeLaGrid ;
{$ENDIF}
end;

Procedure FicheCodeSect(Axe : String ; Var Cod,Lib : String) ;
var
  FCodesect:TFCodesect ;
begin
  if _Blocage(['nrCloture'],False,'nrAucun') then Exit ;
  FCodesect := TFCodesect.Create(Application) ;
  try
    if VH^.Cpta[AxeToFb(Axe)].Structure then begin
      FCodesect.UnAxe := Axe ;
      FCodesect.Cod := Cod ;
      FCodesect.Lib := Lib ;
      FCodesect.ShowModal ;
    end;
  finally
    Cod := FCodesect.Cod ;
    Lib := FCodesect.Lib ;
    FCodesect.Free ;
  end;
  Screen.Cursor := crDefault ;
END ;

procedure TFCodesect.WMGetMinMaxInfo(var MSG: Tmessage);
begin
  with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFCodesect.FormCreate(Sender: TObject);
begin
  WMinX := Width ;
  WMinY := Height ;
end;

procedure TFCodesect.BAideClick(Sender: TObject);
begin
  CallHelpTopic(Self) ;
end;

end.
