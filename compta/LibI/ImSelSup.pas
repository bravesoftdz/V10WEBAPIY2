unit ImSelSup;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, HSysMenu, Menus, Db, DBTables, Hqry, ComCtrls, HRichEdt, ExtCtrls,
  Grids, DBGrids, HDB, HTB97, StdCtrls, ColMemo, Hctrls, HEnt1, HRichOLE,
  HPanel, UIUtil,hstatus ;

type
  TSelSuppImmo = class(TFMul)
    bConsulte: TToolbarButton97;
    QIL_IMMO: TStringField;
    QIL_CPTEMUTATION: TStringField;
    QIL_LIBELLE: TStringField;
    QIL_TYPEMODIF: TStringField;
    QIL_BLOCNOTE: TMemoField;
    QIL_DATEOP: TDateTimeField;
    QIL_TYPEOP: TStringField;
    QIL_MOTIFCES: TStringField;
    QIL_MONTANTCES: TFloatField;
    QIL_MONTANTECL: TFloatField;
    QIL_MONTANTDOT: TFloatField;
    QIL_QTECEDEE: TFloatField;
    QIL_QTEECLAT: TFloatField;
    QIL_DUREE: TIntegerField;
    QIL_METHODE: TStringField;
    QIL_CALCCESSION: TStringField;
    QIL_VOCEDEE: TFloatField;
    QIL_TVAAREVERSER: TFloatField;
    QIL_MONTANTREI: TFloatField;
    QIL_ORDRE: TIntegerField;
    QIL_CODEMUTATION: TStringField;
    QIL_METHODEECO: TStringField;
    QIL_DUREEECO: TIntegerField;
    QIL_METHODEFISC: TStringField;
    QIL_DUREEFISC: TIntegerField;
    QIL_REVISIONDOTECO: TFloatField;
    QIL_REVISIONREPECO: TFloatField;
    QIL_REVISIONREPFISC: TFloatField;
    QIL_REVISIONDOTFISC: TFloatField;
    QIL_TYPEDOT: TStringField;
    QIL_PLANACTIFAV: TIntegerField;
    QIL_PLANACTIFAP: TIntegerField;
    QIL_PVALUE: TFloatField;
    //QIL_VNC: TFloatField;
    QIL_CODEECLAT: TStringField;
    QIL_REVISIONECO: TFloatField;
    QIL_REVISIONFISC: TFloatField;
    QIL_CODECB: TStringField;
    QIL_TVARECUPEREE: TFloatField;
    QIL_TVARECUPERABLE: TFloatField;
    QIL_REPRISEECO: TFloatField;
    QIL_REPRISEFISC: TFloatField;
    QIL_ORDRESERIE: TIntegerField;
    QIL_ETABLISSEMENT: TStringField;
    QIL_LIEUGEO: TStringField;
    QIL_TYPEEXC: TStringField;
    QIL_MONTANTEXC: TFloatField;
    QIL_MONTANTAVMB: TFloatField;
    QIL_BASEECOAVMB: TFloatField;
    QIL_BASEFISCAVMB: TFloatField;
    QIL_DOTCESSECO: TFloatField;
    QIL_DOTCESSFIS: TFloatField;
    QIL_CUMANTCESECO: TFloatField;
    QIL_CUMANTCESFIS: TFloatField;
    procedure BOuvrirClick(Sender: TObject); override ;
    procedure BAnnulerClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BReduireClick(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
    requete: string ;
    QLog: TQuery ;
//    procedure AffecteImmoASupprimer(sRequete : string);
//    function AfficheImmoASupprimer : TModalResult;
//    procedure DestructionImmoASupprimer;
//    function RecupereImmoASupprimer : THQuery;
//    procedure InitialiseListeASupprimer;
//    function RecupereListeASupprimer : THDBGrid;
  end;

implementation

uses imogen;

{$R *.DFM}

procedure CreationImmoASupprimer(Inside : THPanel ; Critere : String; filtrerSur : string) ;
var X : TSelSuppImmo ;
begin
X:=TSelSuppImmo.Create(Application) ;
//x.Q:=OpenSql('
//X.Q.Liste:='MULLIGNE' ;
if Inside=nil then
  begin
  try
    X.ShowModal ;
    finally
    X.Free ;
    end ;
  end
else
  begin
  InitInside(X,Inside) ;
  X.Show ;
  end ;
SourisNormale ;
end ;


(*procedure TSelSuppImmo.AffecteImmoASupprimer(sRequete : string);
begin
  Q.SQL.Clear;
  Q.SQL.Add(sRequete);
end;

procedure TSelSuppImmo.InitialiseListeASupprimer;
begin
if FListe.AllSelected then FListe.AllSelected:=false
                      else FListe.ClearSelected;
end;  *)

procedure TSelSuppImmo.BOuvrirClick(Sender: TObject);
var i: integer ;
begin
inherited;
if FListe.AllSelected then
  begin
  InitMove(100,'');
  Q.First;
  while not Q.EOF do
    begin
    MoveCur(False);
    (*TraiteEnregLog(Q);
    if (fTypeOpe='CES') or (fTypeOpe='CEP') then AnnuleCession else EnregAnnuleOp ;*)
    Q.Delete ;
    Q.Next;
    end;
  FiniMove;
  end
else
  begin
  InitMove(FListe.nbSelected,'');
  for i:=0 to FListe.nbSelected-1 do
    begin
    MoveCur(False);
    FListe.GotoLeBookmark(i);
    (*TraiteEnregLog(QSupp);
    if (fTypeOpe='CES') or (fTypeOpe='CEP') then AnnuleCession else EnregAnnuleOp ;*)
    Q.Delete ;
    end;
  FiniMove;
  end;

end;

procedure TSelSuppImmo.BAnnulerClick(Sender: TObject);
begin
  inherited;
  ModalResult:=mrCancel;
end;

procedure TSelSuppImmo.FormShow(Sender: TObject);
begin
  inherited;
  Q.Active:=true;
end;

procedure TSelSuppImmo.BReduireClick(Sender: TObject);
begin
FicheImmobilisation(nil,Q.FindField('IL_IMMO').AsString,taConsult,'') ;
end;

end.
