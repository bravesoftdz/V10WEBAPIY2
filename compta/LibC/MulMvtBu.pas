{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 25/01/2005
Modifié le ... :   /  /    
Description .. : Remplacé en eAGL par MULMVTBU_TOF.PAS
Mots clefs ... : 
*****************************************************************}
unit MulMvtBu;

interface      

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, HSysMenu, Menus, DB, DBTables, Hqry, StdCtrls, Grids, DBGrids, HDB,
  ComCtrls, HRichEdt, Hctrls, ExtCtrls, Buttons, Hcompte, Mask, hmsgbox,
  Ent1, Hent1, SaisBud, SaisUtil, HTB97, ColMemo, HPanel, UiUtil, HRichOLE ;

Procedure MultiCritereMvtBud(Comment : TActionFiche ; QuelleSaisie : String) ;

type
  TFMulMvtBu = class(TFMul)
    HM: THMsgBox;
    BE_BUDJAL: THValComboBox;
    TBE_BUDJAL: TLabel;
    TBE_NUMEROPIECE: TLabel;
    BE_NUMEROPIECE: THCritMaskEdit;
    BE_NUMEROPIECE_: THCritMaskEdit;
    TBE_NUMEROPIECE_: TLabel;
    BE_DATECOMPTABLE: THCritMaskEdit;
    BE_DATECOMPTABLE_: THCritMaskEdit;
    TBE_DATECOMPTABLE2: THLabel;
    TBE_DATECOMPTABLE: THLabel;
    TBE_EXERCICE: THLabel;
    BE_EXERCICE: THValComboBox;
    BE_AXE: THValComboBox;
    TBE_AXE: TLabel;
    TBE_NATURE: TLabel;
    BE_NATUREBUD: THValComboBox;
    TBE_QUALIFPIECE: THLabel;
    BE_QUALIFPIECE: THValComboBox;
    TBE_UTILISATEUR: THLabel;
    BE_UTILISATEUR: THValComboBox;
    TBE_ETABLISSEMENT: TLabel;
    BE_ETABLISSEMENT: THValComboBox;
    BE_REFINTERNE: TEdit;
    TBE_REFINTERNE: THLabel;
    TBE_GENERAL: THLabel;
    BE_BUDGENE: THCpteEdit;
    BE_VALIDE: TCheckBox;
    Pzlibre: TTabSheet;
    Bevel5: TBevel;
    TBE_TABLE0: TLabel;
    TBE_TABLE2: TLabel;
    BE_TABLE2: THCpteEdit;
    BE_TABLE0: THCpteEdit;
    TBE_TABLE1: TLabel;
    TBE_TABLE3: TLabel;
    BE_TABLE3: THCpteEdit;
    BE_TABLE1: THCpteEdit;
    procedure FormShow(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure BE_EXERCICEChange(Sender: TObject);
    procedure QBeforeOpen(DataSet: TDataSet);
  private
  public
    TypeAction : TActionFiche ;
  end;


implementation

{$R *.DFM}

Procedure MultiCritereMvtBud(Comment : TActionFiche ; QuelleSaisie : String) ;
var FMulMvtBu : TFMulMvtBu ;
    M : RMVT ;
    PP : THPanel ;
BEGIN
if Comment=taCreat then
   BEGIN
   FillChar(M,Sizeof(M),#0) ;
   M.Etabl:=VH^.EtablisDefaut ; M.CodeD:=V_PGI.DevisePivot ;
   M.Simul:='N' ; M.Nature:='INI' ; M.TypeSaisie:=QuelleSaisie ;
   SaisieBudget(taCreat,M,False) ;
   Exit ;
   END ;
FMulMvtBu:=TFMulMvtBu.Create(Application) ;
FMulMvtBu.Q.Manuel:=TRUE ;
Case Comment Of
  taConsult : begin
              FMulMvtBu.FNomFiltre:='MULMTVBU' ;
              FMulMvtBu.Caption:=FMulMvtBu.HM.Mess[0] ;
              FMulMvtBu.Q.Liste:='MULVECRBUD' ;
              end ;
  taModif   : begin
              FMulMvtBu.FNomFiltre:='MULMTMBU' ;
              FMulMvtBu.Caption:=FMulMvtBu.HM.Mess[1] ;
              FMulMvtBu.Q.Liste:='MULMECRBUD' ;
              end ;
  end ;
FMulMvtBu.TypeAction:=Comment ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FMulMvtBu.ShowModal ;
    finally
     FMulMvtBu.Free ;
    end;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FMulMvtBu,PP) ;
   FMulMvtBu.Show ;
   END ;
END ;

procedure TFMulMvtBu.FormShow(Sender: TObject);
begin
BE_DATECOMPTABLE.text:=StDate1900 ; BE_DATECOMPTABLE_.text:=StDate2099 ;
LibellesTableLibre(PzLibre,'TBE_TABLE','BE_TABLE','U') ;
PositionneEtabUser(BE_ETABLISSEMENT) ;
  inherited;
Case TypeAction Of
  taConsult : HelpContext:=15221000 ;
  taModif   : BEGIN HelpContext:=15230000 ; PCritere.HelpContext:=15230000 ; END ; 
  End ; 
end;

procedure TFMulMvtBu.FListeDblClick(Sender: TObject);
begin
  inherited;
if (Q.Eof) And (Q.Bof) then Exit ;
TrouveEtLanceSaisBud(Q,TypeAction) ;
end;

procedure TFMulMvtBu.BE_EXERCICEChange(Sender: TObject);
begin
  inherited;
ExoToDates(BE_EXERCICE.Value,BE_DATECOMPTABLE,BE_DATECOMPTABLE_) ;
end;

procedure TFMulMvtBu.QBeforeOpen(DataSet: TDataSet);
begin
  inherited;
  Q.SQL.Text := ModifieRequeteGrille(Q.SQL.Text); {FP FQ16008-FQ16048}
end;
end.
