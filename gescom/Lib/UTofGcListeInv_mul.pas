{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 14/12/2000
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : GCLISTEINV_MUL ()
Mots clefs ... : TOF;GCLISTEINV_MUL
*****************************************************************}
Unit UTOFGCLISTEINV_MUL ;

Interface

Uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
{$IFDEF EAGLCLIENT}
      eMul,MaineAGL,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}db,Mul,FE_Main,
{$ENDIF}
     HCtrls, HEnt1, HMsgBox, HTB97, UTOF, UTOB ;


function GCLanceFiche_ListeInvMul(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Type
  TOF_GCLISTEINV_MUL = Class (TOF)
  private
    EnContremarque : boolean;
    TOBSelection : TOB;
    G_Inven: THGRID ;
    BSUPPRIMER : TToolbarButton97;
    function  PutSelectionIntoTOB : boolean;
    procedure BSUPPRIMERClick(Sender: TObject);
  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  end ;

const FMess : Array[0..1] of string = ('Veuillez sélectionner au moins une liste à supprimer',
                                       '0;?caption?;Confirmez vous la suppression ?;Q;YN;Y;N;');

Implementation

function GCLanceFiche_ListeInvMul(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
Result:='';
if Nat='' then exit;
if Cod='' then exit;
Result:=AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

{==============================================================================================}
{================================== Procédure de la TOF =======================================}
{==============================================================================================}
procedure TOF_GCLISTEINV_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_GCLISTEINV_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_GCLISTEINV_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_GCLISTEINV_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_GCLISTEINV_MUL.OnArgument (S : String ) ;
begin
  Inherited ;
EnContremarque := (pos('CONTREMARQUE',S) > 0);
G_Inven:=THGRID(GetControl('FListe'));
BSUPPRIMER:=TToolbarButton97(GetControl('BOUVRIR'));
BSUPPRIMER.OnClick:=BSUPPRIMERCLick;
end ;

procedure TOF_GCLISTEINV_MUL.OnClose ;
begin
  Inherited ;
end ;

//
//-------------------------------------------------------------------
//
function TOF_GCLISTEINV_MUL.PutSelectionIntoTOB : boolean;
var i : integer;
begin
result := true;
with TFMul(Ecran) do
  begin
  if (FListe.NbSelected = 0) and (not FListe.AllSelected) then
    begin
    PGIInfo(FMess[0], Caption);
    result := false;
    exit;
    end;

  TOBSelection := TOB.Create('', nil, -1);
{$IFDEF EAGLCLIENT}
  if FListe.AllSelected then
    TOBSelection.LoadDetailDB('LISTEINVENT', '', '', Q.TQ,false)
{$ELSE}
  Q.DisableControls;
  if FListe.AllSelected then
    TOBSelection.LoadDetailDB('LISTEINVENT', '', '', Q,false)
{$ENDIF}
   else
    begin
    for i := 0 to FListe.NbSelected-1 do
      begin
      FListe.GotoLeBookMark(i);
{$IFDEF EAGLCLIENT}
      Q.TQ.Seek(FListe.Row-1) ;
      TOB.Create('LISTEINVENT', TOBSelection, -1).SelectDB('',Q.TQ, true);
{$ELSE}
      TOB.Create('LISTEINVENT', TOBSelection, -1).SelectDB('', Q, true);
{$ENDIF}
      end;
    end;
{$IFDEF EAGLCLIENT}
{$ELSE}
  Q.EnableControls;
{$ENDIF}
  end;
end;

procedure TOF_GCLISTEINV_MUL.BSUPPRIMERClick(Sender: TObject);
var
    i_ind1 : integer;
    TobDel : TOB;
    SQL : string;

begin
if not PutSelectionIntoTOB then exit;
if HShowMessage(FMess[1], Ecran.Caption, '') = mrYes then
    for i_ind1 := 0 to TobSelection.Detail.Count - 1 do
        begin
        TobDel := TobSelection.Detail[i_ind1];
        if  EnContremarque then
            begin
            SQL := 'Delete from LISTEINVLIGCONTREM where GIM_CODELISTE="' + TobDel.GetValue('GIE_CODELISTE') +
               '"';
            ExecuteSQL(SQL);
            end else
            begin
            SQL := 'Delete from LISTEINVLIG where GIL_CODELISTE="' + TobDel.GetValue('GIE_CODELISTE') +
                   '"';
            ExecuteSQL(SQL);

            SQL := 'Delete from LISTEINVLOT where GLI_CODELISTE="' + TobDel.GetValue('GIE_CODELISTE') +
                   '"';
            ExecuteSQL(SQL);
            end;
        TobDel.DeleteDB;
        end;
TFMul(Ecran).bCherche.Click;
end;

Initialization
  registerclasses ( [ TOF_GCLISTEINV_MUL ] ) ;
end.
