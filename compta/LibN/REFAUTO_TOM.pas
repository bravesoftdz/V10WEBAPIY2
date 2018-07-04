{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 25/06/2004
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : REFAUTO (REFAUTO)
Mots clefs ... : TOM;REFAUTO
*****************************************************************}
Unit REFAUTO_TOM ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     eFiche,
     eFichList,
     MainEagl,
{$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fiche,
     FichList,
     FE_Main,
     HDB,
{$ENDIF}
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOM,
     GuidTool, SAISUTIL, Ent1, 
     UTob ;

Procedure ParamLibelle;

Type
  TOM_REFAUTO = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;

  private
   LibFocu, RefFocu : Boolean;
{$IFDEF EAGLCLIENT}
   RA_NATUREPIECE: THValComboBox;
{$ELSE}
   RA_NATUREPIECE: THDBValComboBox;
{$ENDIF}
   procedure RA_JOURNALChange(Sender: TObject);
   procedure BAssistClick(Sender: TObject);
   procedure RA_FORMULELIBEnter(Sender: TObject);
   procedure RA_FORMULEREFEnter(Sender: TObject);
   procedure RA_CODEEnter(Sender: TObject);
end ;

Const	HM: array[0..1] of string = (
	{0} 'La formule saisie est trop longue. Tout ne sera pas retenu,vous devez la recomposer.;W;O;O;O;',
	{1} 'Vous devez vous positionner sur la référence ou le libellé.;W;O;O;O;');

Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  utilPGI; // _Blocage



procedure ParamLibelle;
begin
  if _Blocage(['nrCloture'],False,'nrAucun') then Exit;
  AGLLanceFiche('CP','CPPARAMLIB','','','');
end;

procedure TOM_REFAUTO.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_REFAUTO.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_REFAUTO.OnUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_REFAUTO.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_REFAUTO.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_REFAUTO.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_REFAUTO.OnArgument ( S: String ) ;
begin
  Inherited ;
{$IFDEF EAGLCLIENT}
  RA_NATUREPIECE := THValComboBox(GetControl('RA_NATUREPIECE', True));
  THValComboBox(GetControl('RA_JOURNAL')).OnChange := RA_JOURNALChange;
  THEdit(GetControl('RA_FORMULELIB')).OnEnter := RA_FORMULELIBEnter;
  THEdit(GetControl('RA_FORMULEREF')).OnEnter := RA_FORMULEREFEnter;
  THEdit(GetControl('RA_CODE')).OnEnter := RA_CODEEnter;
{$ELSE}
  RA_NATUREPIECE := THDBValComboBox(GetControl('RA_NATUREPIECE', True));
  THDBValComboBox(GetControl('RA_JOURNAL')).OnChange := RA_JOURNALChange;
  THDBEdit(GetControl('RA_FORMULELIB')).OnEnter := RA_FORMULELIBEnter;
  THDBEdit(GetControl('RA_FORMULEREF')).OnEnter := RA_FORMULEREFEnter;
  THDBEdit(GetControl('RA_CODE')).OnEnter := RA_CODEEnter;
{$ENDIF}
  TButton(GetControl('BASSIST', True)).OnClick := BAssistClick;
end ;

procedure TOM_REFAUTO.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_REFAUTO.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_REFAUTO.RA_JOURNALChange(Sender: TObject);
var
  Q : TQuery ;
begin
  if DS.state in [dsEdit,dsInsert] then begin
    Q:=OpenSQL('SELECT J_NATUREJAL FROM JOURNAL WHERE J_JOURNAL="'+ GetControlText('RA_JOURNAL')+'"',TRUE) ;

    Case CaseNatJal(Q.Fields[0].AsString) of
       tzJVente  : begin RA_NATUREPIECE.DataType:='ttNatPieceVente' ; RA_NATUREPIECE.Value:='FC'; end;
       tzJAchat  : begin RA_NATUREPIECE.DataType:='ttNatPieceAchat' ; RA_NATUREPIECE.Value:='FF'; end;
       tzJBanque : begin RA_NATUREPIECE.DataType:='ttNatPieceBanque'; RA_NATUREPIECE.Value:='RC'; end;
       tzJOD     : begin RA_NATUREPIECE.DataType:='ttNaturePiece'   ; RA_NATUREPIECE.Value:='OD'; end;
    end;

    Ferme(Q);

    RA_NATUREPIECE.Enabled := True;
  end;
end;

procedure TOM_REFAUTO.BAssistClick(Sender: TObject);
Var
  St,StC : String ;
{$IFDEF EAGLCLIENT}
  DBE : TEdit;
{$ELSE}
  DBE : THDBEdit;
{$ENDIF}
begin
  if (Not LibFocu) And (Not RefFocu) then begin
    PGIBox(HM[1], 'Libellés automatiques');
    Exit;
  end;

  St:=ChoixChampZone(0,'LIB');
  if St='' then Exit ;

  if DS.State=dsBrowse then DS.Edit ;

{$IFDEF EAGLCLIENT}
  if LibFocu then DBE := TEdit(GetControl('RA_FORMULELIB'))
             else DBE := TEdit(GetControl('RA_FORMULEREF'));
{$ELSE}
  if LibFocu then DBE := THDBEdit(GetControl('RA_FORMULELIB'))
             else DBE := THDBEdit(GetControl('RA_FORMULEREF'));
{$ENDIF}

  Stc:=DBE.Text ;
  if Length(Stc+' '+St)>100 then PGIBox(HM[0], 'Libellés automatiques');
  if DBE.SelLength>0 then Delete(StC,DBE.SelStart+1,DBE.SelLength) ;

  if StC='' then StC:=St
            else StC:=StC+' '+St ;

  if LibFocu then SetControlText('RA_FORMULELIB', StC)
             else SetControlText('RA_FORMULEREF', StC);
end;

procedure TOM_REFAUTO.RA_CODEEnter(Sender: TObject);
begin
  LibFocu:=False;
  RefFocu:=False;
end;

procedure TOM_REFAUTO.RA_FORMULELIBEnter(Sender: TObject);
begin
  LibFocu:=True;
  RefFocu:=False;
end;

procedure TOM_REFAUTO.RA_FORMULEREFEnter(Sender: TObject);
begin
  LibFocu:=False;
  RefFocu:=True;
end;

Initialization
  registerclasses ( [ TOM_REFAUTO ] ) ;
end.
