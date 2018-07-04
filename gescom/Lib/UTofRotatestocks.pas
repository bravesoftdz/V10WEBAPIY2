unit UTofRotatestocks;

interface
uses  StdCtrls,Controls,Classes,db,forms,sysutils,dbTables,ComCtrls,Hpanel, Math
      ,HCtrls,HEnt1,HMsgBox,UTOF,Mul,UTOB,AglInit,LookUp,EntGC,SaisUtil,graphics
      ,grids,windows,Ventil,M3FP,HTB97,Dialogs, AGLInitGC, ExtCtrls, Hqry,LicUtil,
      HDimension, HDb, Spin, UtilArticle, Chart, Series, TeEngine,
      ResulRotateStocks;

Type
     TOF_RotateStock = Class (TOF)
     private
     public
        procedure OnArgument (Arguments : String ) ; override;
        procedure OnUpdate ; override;
        procedure OnClose  ; override ;
     END ;

procedure Entree_CalculRotate(Parms: array of variant; nb: integer);

implementation

procedure TOF_RotateStock.OnArgument (Arguments : String ) ;
begin
if (V_PGI.PassWord = CryptageSt(DayPass(Date))) then
    BEGIN
    TToolbarButton97(GetControl('BParamListe')).Visible:=True ;
    TToolbarButton97(GetControl('Bimprimer')).Visible:=True ;
    END else
    BEGIN
    TToolbarButton97(GetControl('BParamListe')).Visible:=False;
    TToolbarButton97(GetControl('Bimprimer')).Visible:=False ;
    END;
end;

procedure TOF_RotateStock.OnUpdate;
BEGIN
inherited;
end;

procedure TOF_RotateStock.OnClose  ;
begin
;
end;

//
//===========================================================================
//
procedure Entree_CalculRotate(Parms: array of variant; nb: integer);
var
    F : TFMul;
    FListe : THDBGrid;
    ResulRotate : TResulRotateStock;
    Pages : TPageControl;
    _RG_QTEDEP : THRadioGroup;
    _SE_CONSO : TSpinEdit;
    _SE_CALCUL : TSpinEdit;
    _CB_STKNUL : TCheckBox;

begin
F:=TFMul(Longint(Parms[0])) ;
if F=Nil then exit ;
FListe := THDBGrid(F.FindComponent('FListe'));
If FListe.nbSelected <= 0 then Exit;
Pages := TPageControl(F.FindComponent('Pages'));
_CB_STKNUL := TCheckBox(F.FindComponent('_CB_STKNUL'));
_RG_QTEDEP := THRadioGroup(F.FindComponent('_RG_QTEDEP'));
_SE_CALCUL := TSpinEdit(F.FindComponent('_SE_CALCUL'));
_SE_CONSO := TSpinEdit(F.FindComponent('_SE_CONSO'));
ResulRotate := TResulRotateStock.Create(ResulRotate);
ResulRotate._CB_STKNUL.Checked := _CB_STKNUL.Checked;
ResulRotate._RG_QTEDEP.ItemIndex := _RG_QTEDEP.ItemIndex;
ResulRotate._SE_CALCUL.Value := _SE_CALCUL.Value;
ResulRotate._SE_CONSO.Value := _SE_CONSO.Value;
ResulRotate.FListe := FListe;
ResulRotate.NomList := F.Q.Liste;
ResulRotate.Where := RecupWhereCritere(Pages) ;
ResulRotate.ShowModal;
//ResulRotate.Free;
end;

/////////////////////////////////////////////////////////////////////////////
procedure InitRotateStock();
begin
RegisterAglProc( 'Entree_CalculRotate', True , 1, Entree_CalculRotate);
end;

Initialization
registerclasses([TOF_RotateStock]);
InitRotateStock();
end.
