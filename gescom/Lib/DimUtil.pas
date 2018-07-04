unit DimUtil;

interface
uses {$IFDEF VER150} variants,{$ENDIF} HCtrls, HPanel, UTOB, EntGC ;

function MyRechDom(FTypeTable, Value, Defaut : String; Abrege : boolean) : string;
Procedure DisplayDimensions(ZeTOB : TOB; Panel : THPanel; LabelBaseName, EditBaseName : String; PremDim:integer; ShowHidePanel : Boolean);

implementation
uses SysUtils,
{$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     InvUtil, HDimension;

{***********A.G.L.***********************************************
Auteur  ...... : TG
Créé le ...... : 09/02/2000
Modifié le ... :   /  /
Description .. : identique à RechDom sans retourner "Error" si Value est vide
Mots clefs ... : RECHDOM;TABLETTE
*****************************************************************}
function MyRechDom(FTypeTable, Value, Defaut : String; Abrege : boolean) : string;
begin
result := Value;
if result <> '' then result := RechDom(FTypeTable, Value, Abrege);
if result = '' then result := Defaut;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TG
Créé le ...... : 18/02/2000
Modifié le ... :   /  /
Description .. : Affiche les dimensions extraites d'une TOB sur un panel contenant 5 labels et 5 edit
Mots clefs ... : DIMENSION
*****************************************************************}
Procedure DisplayDimensions(ZeTOB : TOB; Panel : THPanel; LabelBaseName, EditBaseName : String; PremDim:integer; ShowHidePanel : Boolean);
var TLIB : THLabel;
    CHPS : THCritMaskEdit;
    i_ind, i_dim : integer;
    b_dim : boolean;
    GrilleDim,CodeDim : String ;
begin
i_dim := PremDim;
b_dim := false;
for i_ind := PremDim to MaxDimension do
  begin
  TLIB := THLabel(Panel.Owner.FindComponent(LabelBaseName + IntToStr(i_dim)));
  CHPS := THCritMaskEdit(Panel.Owner.FindComponent(EditBaseName + IntToStr(i_dim)));
  GrilleDim:=VarToStr(ZeTOB.GetValue('GA_GRILLEDIM'+IntToStr(i_ind))) ;
  CodeDim:=VarToStr(ZeTOB.GetValue('GA_CODEDIM'+IntToStr(i_ind))) ;
  if GrilleDim<>'' then
     begin
     TLIB.Caption:=RechDom('GCGRILLEDIM'+IntToStr(i_ind),GrilleDim,False) ;
     CHPS.Text:=GCGetCodeDim(GrilleDim,CodeDim,i_ind) ;
     CHPS.Visible:=True ; b_dim:=True ;
     inc(i_dim);
     end;
  end;
for i_ind:=i_dim to 5 do
  begin
  THLabel(Panel.Owner.FindComponent(LabelBaseName + IntToStr(i_ind))).Caption := '';
  THCritMaskEdit(Panel.Owner.FindComponent(EditBaseName + IntToStr(i_ind))).Visible := false;
  end;

if ShowHidePanel then Panel.Visible := b_dim;
end;


end.
 
