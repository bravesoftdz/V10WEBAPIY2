unit profiluser;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  FichList, DB, DBTables, StdCtrls, Hctrls, Mask, DBCtrls, hmsgbox, Region,
  Buttons, ExtCtrls, Grids, DBGrids, HDB, HEnt1, Ent1, CodePost, HSysMenu,
  Hqry, HTB97, HPanel, UiUtil ;

Procedure FicheProfilUser(Mode : TActionFiche) ;

type
  TFProfilUser = class(TFFicheListe)
    CPU_COMPTE1: TDBEdit;
    TPUC_COMPTE1: THLabel;
    TPUC_USERGRP: THLabel;
    CPU_USERGRP: THDBValComboBox;
    Bevel1: TBevel;
    TPUC_USER: THLabel;
    CPU_USER: THDBValComboBox;
    TPUC_EXCLUSION1: THLabel;
    CPU_EXCLUSION1: TDBEdit;
    TPUC_COMPTE2: THLabel;
    CPU_COMPTE2: TDBEdit;
    TPUC_EXCLUSION2: THLabel;
    CPU_EXCLUSION2: TDBEdit;
    TPUC_TABLELIBRE1: THLabel;
    CPU_TABLELIBRE1: TDBEdit;
    TPUC_TABLELIBRE2: THLabel;
    CPU_TABLELIBRE2: TDBEdit;
    TPUC_TYPE: THLabel;
    CPU_TYPE: THDBValComboBox;
    TaCPU_USERGRP: TStringField;
    TaCPU_USER: TStringField;
    TaCPU_TYPE: TStringField;
    TaCPU_COMPTE1: TStringField;
    TaCPU_EXCLUSION1: TStringField;
    TaCPU_COMPTE2: TStringField;
    TaCPU_EXCLUSION2: TStringField;
    TaCPU_TABLELIBRE1: TStringField;
    TaCPU_TABLELIBRE2: TStringField;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure CPU_TABLELIBRE1DblClick(Sender: TObject);
    procedure CPU_TABLELIBRE2DblClick(Sender: TObject);
  private    { Déclarations privées }
    Function EnregOK : boolean ; Override ;
    Procedure NewEnreg ; Override ;
  public    { Déclarations publiques }
  end;

implementation

uses Devise, TabLiRub;

{$R *.DFM}

Procedure FicheProfilUser(Mode : TActionFiche) ;
var XX : TFProfilUser ;
    PP       : THPanel ;
BEGIN
if Blocage(['nrCloture','nrBatch'],True,'nrBatch') then Exit ;
XX:=TFProfilUser.Create(Application) ;
XX.InitFL('CPU','PRT_ETABLIS','','',Mode,TRUE,Nil,Nil,Nil,[]) ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
   Try
    XX.ShowModal ;
   Finally
    XX.Free ;
    Bloqueur('nrBatch',False) ;
   End ;
   Screen.Cursor:=crDefault ;
   END else
   BEGIN
   InitInside(XX,PP) ;
   XX.Show ;
   END ;
END ;

procedure TFProfilUser.FormShow(Sender: TObject);
begin
  inherited;
if (Ta.Eof) And (Ta.Bof) And (FTypeAction<>taConsult) then
   BEGIN
   if ta.State=dsInsert then NewEnreg else BinsertClick(Nil) ;
   END ;
end;

procedure TFProfilUser.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if isInside(Self) then Bloqueur('nrBatch',False) ;
  inherited;
ChargeProfilUser ;
end;

Function TFProfilUser.EnregOK : boolean ;
BEGIN
result:=Inherited EnregOK  ; if Not Result then Exit ;
if ((TaCPU_USERGRP.AsString='') or (CPU_USERGRP.Value='')) then TaCPU_USERGRP.AsString:='...' ;
if ((TaCPU_USER.AsString='') or (CPU_USER.Value='')) then TaCPU_USER.AsString:='...' ;
if Ta.State=dsinsert then
   BEGIN
   if PresenceComplexe('CPPROFILUSERC',
                       ['CPU_USERGRP','CPU_USER','CPU_TYPE'],
                       ['=','=','='],
                       [TaCPU_USERGRP.AsString,TaCPU_USER.AsString,TaCPU_TYPE.AsString],
                       ['S','S','S']) then BEGIN Result:=FALSE ; Exit ; END ;
   END ;
END ;

Procedure TFProfilUser.NewEnreg ;
BEGIN
Inherited ;
TaCPU_USERGRP.AsString:='...' ;
TaCPU_USER.AsString:='...' ;
END ;

Procedure TransFormeCode(Var St : String ; PVirgule : Boolean) ;
BEGIN
if St='' then Exit ;
if PVirgule then While Pos(',',St)>0 do St[Pos(',',St)]:=';'
            else While Pos(';',St)>0 do St[Pos(';',St)]:=',' ;
if PVirgule then St:=St+';' else Delete(St,Length(St),1) ;
END ;

Function Min(a,b : Integer) : Integer ;
BEGIN if a>b then Result:=b else Result:=a ; END ;

Function ZoomTL(fb : tFichierBase ; Cpt : String) : String ;
Var CodDe,CodA,Stock,St : String ;
    i,j : Integer ;
begin
CodDe:='' ; CodA:='' ; Stock:='' ;
If Cpt<>'' Then
  BEGIN
  i:=Pos(':',Cpt) ;
  if i>0 then
     BEGIN
     CodDe:=Copy(Cpt,1,i-1) ;
     CodA:=Copy(Cpt,i+1,Length(CodDe)) ;
     Stock:=Copy(Cpt,2*Length(CodDe)+2,Length(Cpt)) ;
     TransformeCode(CodDe,True) ; TransformeCode(CodA,True) ;
     END else
     BEGIN
     i:=Pos('&',Cpt) ; j:=Pos('|',Cpt) ;
     if (i=0) And (j=0) then CodDe:=Cpt ;
     if (i<>0) And (j=0) then BEGIN CodDe:=Copy(Cpt,1,i-2) ; Stock:=Copy(Cpt,i,Length(Cpt)) ; END ;
     if (j<>0) And (i=0) then BEGIN CodDe:=Copy(Cpt,1,j-2) ; Stock:=Copy(Cpt,j,Length(Cpt)) ; END ;
     if (i<>0) And (j<>0)then BEGIN CodDe:=Copy(Cpt,1,Min(i,j)-2) ; Stock:=Copy(Cpt,Min(i,j),Length(Cpt)) ; END ;
     TransformeCode(CodDe,True) ; CodA:=CodDe ;
     END ;
  END ;
if Stock<>'' then
   if Stock[1]=',' then Stock:=Copy(Stock,2,Length(Stock)) ;
ChoixTableLibrePourRub(fb,'',Stock,CodDe,CodA) ;
TransformeCode(CodDe,False) ; TransformeCode(CodA,False) ;
if CodA<>CodDe then St:=CodDe+':'+CodA else St:=CodDe ;
if Stock<>'' then BEGIN if Stock[1]<>',' then Stock:=','+Stock ; St:=St+Stock ; END ;
Result:=St ;
end;

procedure TFProfilUser.CPU_TABLELIBRE1DblClick(Sender: TObject);
Var Cpt : String ;
begin
  inherited;
Cpt:=trim(CPU_TableLibre1.Text) ;
If ta.State=dsBrowse Then Ta.Edit ;
CPU_TableLibre1.Text:=ZoomTL(fbGene,Cpt) ;
end;

procedure TFProfilUser.CPU_TABLELIBRE2DblClick(Sender: TObject);
Var Cpt : String ;
begin
  inherited;
Cpt:=trim(CPU_TableLibre2.Text) ;
If ta.State=dsBrowse Then Ta.Edit ;
CPU_TableLibre2.Text:=ZoomTL(fbAux,Cpt) ;
end;

end.
