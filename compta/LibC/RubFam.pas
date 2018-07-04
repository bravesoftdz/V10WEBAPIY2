unit RubFam;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Tablette, HSysMenu, hmsgbox, Db, DBTables, StdCtrls, DBCtrls, Buttons,
  ExtCtrls, Grids, DBGrids, HDB, Ent1, HEnt1, hCtrls,
  //XMG 05/04/04 début
  {$IFDEF ESP}
  FAMRUB_TOF, 
  {$ELSE}
  FamRub,
  {$ENDIF ESP}
  //XMG 05/04/04 fin
  Hqry, HTB97, HPanel, UiUtil
  ;

procedure FamillesRub ;

type
  TFRubFam = class(TFTablette)
    Brubfam: TToolbarButton97;
    Msg: THMsgBox;
    TChoixCodYDS_TYPE: TStringField;
    TChoixCodYDS_CODE: TStringField;
    TChoixCodYDS_LIBELLE: TStringField;
    TChoixCodYDS_ABREGE: TStringField;
    TChoixCodYDS_LIBRE: TStringField;
    TChoixCodYDS_NODOSSIER: TStringField;
    TChoixCodYDS_PREDEFINI: TStringField;
    procedure BrubfamClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure TChoixCodNewRecord(DataSet: TDataSet);
    procedure BValiderClick(Sender: TObject);
    Function Supprime : Boolean ; override ;
  private { Déclarations privées }
    Function SupprimeFamille : Boolean ;
  public  { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Uses UtilPgi ;

procedure FamillesRub ;
var X  : TFRubFam ;
    PP : THPanel ;
begin
if _Blocage(['nrCloture'],False,'nrAucun') then Exit ;
X:=TFRubFam.Create(Application) ;
X.FQuoi:='TTRUBFAMILLE' ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     X.ShowModal ;
    finally
     X.Free ;
    end ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;
end ;

Function TFRubFam.Supprime : Boolean ;
begin
Result:=SupprimeFamille ;
end ;

Function TFRubFam.SupprimeFamille : Boolean ;
Var Q : TQuery ;
    Trouver : Boolean ;
    St,StLib : String ;
BEGIN
Result:=True ; Trouver:=False ;
Q:=OpenSql('Select RB_FAMILLES from RUBRIQUE',True) ;
While(Not Q.Eof) And (Not Trouver ) do
   BEGIN
   StLib:=Q.Fields[0].AsString ;
   While StLib<>'' do
       BEGIN
       St:=ReadTokenSt(StLib) ;
       if St=TChoixCod.FindField(Fprefixe+'_CODE').AsString then
          BEGIN Trouver:=True ; Break ; END ;
       END ;
   Q.Next ;
   END ;
if Trouver then BEGIN Msg.Execute(0,caption,'') ; Result:=False ; END ;
Ferme(Q) ; Screen.Cursor:=SyncrDefault ;
END ;

procedure TFRubFam.BrubfamClick(Sender: TObject);
Var Famille,LibFam : String ;
    Q : Tquery ;
    Vide : Boolean ;
begin
Q:=OpenSql('Select * from Rubrique',True) ; Vide:=Q.Eof ; Ferme(Q) ;
if Vide then BEGIN Msg.Execute(1,caption,'') ; Exit ; END ;
Famille:='' ; LibFam:='' ;
ParametrageFamilleRubrique('','',Famille,LibFam,False) ;
Screen.Cursor:=SyncrDefault ;
end ;

procedure TFRubFam.BDeleteClick(Sender: TObject);
begin
if Not Supprime then Exit ;
  inherited;
end;

procedure TFRubFam.TChoixCodNewRecord(DataSet: TDataSet);
begin
  inherited;
  if ctxStandard in V_PGI.PGIContexte then
  begin
    TChoixCod.FindField('YDS_PREDEFINI').AsString:='STD';
    TChoixCod.FindField('YDS_NODOSSIER').AsString:='000000';
    TChoixCod.FindField('YDS_CODE').AsString:='$';
  end else
  begin
    TChoixCod.FindField('YDS_PREDEFINI').AsString:='DOS';
    TChoixCod.FindField('YDS_NODOSSIER').AsString:=V_PGI.NoDossier ;
  end;
end;

procedure TFRubFam.BValiderClick(Sender: TObject);
var stCode : string;
begin
  stCode := TChoixCod.FindField('YDS_CODE').AsString;
  if (((ctxStandard in V_PGI.PGIContexte) and (stCode[1]<>'$')) or
    ((not (ctxStandard in V_PGI.PGIContexte)) and ((stCode[1]='$') or (stCode[1]='@')))) then
    PGIInfo ('Code famille incorrect.'+#10#13+'Les codes commençant par "@" sont réservés à Cegid.'+
            #10#13+'Les codes commençant par "$" sont réservés aux standards.',Caption)
  else inherited;
end;

end.
