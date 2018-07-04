unit ImMulHis;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, HSysMenu, Menus, Db, DBTables, Hqry, ComCtrls, HRichEdt, ExtCtrls,
  Grids, DBGrids, HDB, HTB97, StdCtrls, ColMemo, Hctrls,HPanel,UiUtil,Hent1,ImEnt,
  Mask,LookUp, HRichOLE ;

type
  TFImMulHis = class(TFMul)
    XX_WHERE: TEdit;
    tI_ETABLISSEMENT: THLabel;
    I_ETABLISSEMENT: THValComboBox;
    HLabel10: THLabel;
    I_COMPTELIE: THCritMaskEdit;
    HLabel4: THLabel;
    HLabel7: THLabel;
    Label10: TLabel;
    Label12: TLabel;
    HLabel8: THLabel;
    HLabel9: THLabel;
    I_DATEAMORT: THCritMaskEdit;
    I_DATEAMORT_: THCritMaskEdit;
    I_DATEPIECEA_: THCritMaskEdit;
    I_DATEPIECEA: THCritMaskEdit;
    FExercice2: THValComboBox;
    HLabel1: THLabel;
    HLabel2: THLabel;
    Nature: THLabel;
    HLabel3: THLabel;
    HLabel5: THLabel;
    HLabel11: THLabel;
    I_IMMO: TEdit;
    I_LIBELLE: TEdit;
    I_NATUREIMMO: THValComboBox;
    I_QUALIFIMMO: THValComboBox;
    I_LIEUGEO: THValComboBox;
    I_COMPTEIMMO: THCritMaskEdit;
    I_ORGANISMECB: THCritMaskEdit;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure I_COMPTEIMMOElipsisClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject); override ;
    procedure FormShow(Sender: TObject);
    procedure I_ORGANISMECBKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure I_ORGANISMECBElipsisClick(Sender: TObject);
    procedure FormCreate(Sender: TObject); override ;
  private
    { Déclarations privées}
  public
    { Déclarations publiques}
  end;

procedure AfficheHistoriqueImmo;

implementation
uses imogen;

{$R *.DFM}

procedure AfficheHistoriqueImmo;
var FImMulHis: TFImMulHis;
    PP:THPanel;
begin
  FImMulHis :=TFImMulHis.Create(Application) ;
  FImMulHis.FNomFiltre:='MULVIMMOS' ;
  FImMulHis.Q.Liste:='MULVIMMOS' ;
  PP:=FindInsidePanel;
  if PP=nil then
  begin
    try
      FImMulHis.ShowModal ;
    finally
      FImMulHis.Free ;
    end ;
  end else
  begin
    InitInside(FImMulHis,PP);
    FImMulHis.Show;
  end;
  Screen.Cursor:=SyncrDefault ;
end;

procedure TFImMulHis.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
  if IsInside(Self) then Action := caFree;
end;

procedure TFImMulHis.I_COMPTEIMMOElipsisClick(Sender: TObject);
var stWhere : string;
begin
  inherited;
  if THCritMaskEdit(Sender).Name = 'I_COMPTEIMMO' then
  stWhere := '(G_GENERAL>="'+VHImmo^.CpteImmoInf+'" AND G_GENERAL<="'+VHImmo^.CpteImmoSup+'") OR '+
             '(G_GENERAL>="'+VHImmo^.CpteFinInf+'" AND G_GENERAL<="'+VHImmo^.CpteFinSup+'")'
  else if THCritMaskEdit(Sender).Name = 'I_COMPTELIE' then
    stWhere := '(G_GENERAL>="'+VHImmo^.CpteCBInf+'" AND G_GENERAL<="'+VHImmo^.CpteCBSup+'") OR '+
               '(G_GENERAL>="'+VHImmo^.CpteLocInf+'" AND G_GENERAL<="'+VHImmo^.CpteLocSup+'")';
  LookupList(TControl(Sender),'','GENERAUX','G_GENERAL','G_LIBELLE',stWhere,'G_GENERAL', True,0)  ;
end;

procedure TFImMulHis.FListeDblClick(Sender: TObject);
begin
  inherited;
  if(Q.Eof)And(Q.Bof) then Exit ;
  inherited;
  FicheImmobilisation(Q,Q.FindField('I_IMMO').AsString,taConsult,'') ;
end;

procedure TFImMulHis.FormShow(Sender: TObject);
begin
{$IFDEF SERIE1}
  I_ETABLISSEMENT.Visible := False;
  tI_ETABLISSEMENT.Visible := False;
{$ENDIF}
  XX_WHERE.Text:='I_ETAT="FER"';
  I_DATEPIECEA.text:=StDate1900 ; I_DATEPIECEA_.text:=StDate2099 ;
  inherited;
end;

procedure TFImMulHis.I_ORGANISMECBKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
  if Key = VK_F5 then I_ORGANISMECBElipsisClick(Sender) ; //EPZ 25/03/99
end;

procedure TFImMulHis.I_ORGANISMECBElipsisClick(Sender: TObject);
begin
  inherited;
{$IFDEF SERIE1}
LookUpList (TControl (Sender),TraduireMemoire('Auxiliaire'),'TIERS','T_AUXILIAIRE','T_LIBELLE','T_NATUREAUXI="FOU"','T_AUXILIAIRE',True,1) ;
{$ELSE}
LookUpList (TControl (Sender),TraduireMemoire('Auxiliaire'),'TIERS','T_AUXILIAIRE','T_LIBELLE','T_NATUREAUXI="FOU"','T_AUXILIAIRE',True,2) ;
{$ENDIF}
end;

procedure TFImMulHis.FormCreate(Sender: TObject);
begin
  inherited;
{$IFDEF SERIE1}
HelpContext:=511500 ;
{$ELSE}
HelpContext:=2115000 ;
{$ENDIF}

end;

end.
