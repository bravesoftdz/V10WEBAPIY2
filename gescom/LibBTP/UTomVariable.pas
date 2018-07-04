{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 05/09/2003
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : BVARIABLES (BVARIABLES)
Mots clefs ... : TOM;BVARIABLES
*****************************************************************}
Unit UTomvariable ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fiche, 
     FichList,
     HDB,
{$else}
     eFiche, 
     eFichList,
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1,
     HMsgBox, 
     UTOM,
     HTB97,
     UTob ;

Type
  TOM_BVARIABLES = Class (TOM)
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
    TypeVariable  : String;
    //
    Bdelete       : Ttoolbarbutton97;
    BValider      : TtoolBarButton97;
    BInsert       : TToolBarButton97;
    BFerme        : TToolBarButton97;
    //
    FListe        : THDBGrid;
    //
    CodeVariable  : THDBEdit;
    Libelle       : THDBEdit;
    Valeur        : THDBEdit;
    //
    Action        : TActionFiche;
    //
    procedure DeleteOnClick(Sender: Tobject);
    procedure GetEventObjects;
    procedure GetObjects;

    function RechercheVariable(TypeVAR, Article, CodeVar: String; var ValVar: Variant): String;
    procedure RowEnterOnclick(Sender: TObject);

    end ;

Implementation

procedure TOM_BVARIABLES.OnNewRecord ;
begin
  Inherited ;

  if (TypeVariable = 'L') then bvalider.visible := True;

end ;

procedure TOM_BVARIABLES.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_BVARIABLES.OnUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_BVARIABLES.OnAfterUpdateRecord ;
begin
  Inherited ;

  if (TypeVariable = 'L') then bvalider.visible := False;

end ;

procedure TOM_BVARIABLES.OnLoadRecord ;
begin
  Inherited ;

end ;

procedure TOM_BVARIABLES.OnChangeField ( F: TField ) ;
begin
  Inherited ;

  {*
  if F.FieldName = 'BVA_CODEVARIABLE' then
  begin

    CodeVar := GetControlText('BVA_CODEVARIABLE');
    Article := GetControlText('BVA_ARTICLE');

    if CodeVar = '' then exit;

    // Recherche de la valeur dans le cas de l'appel d'une variables de type Générales dans un article
    if not RechercheVariable('A', 'APPLICATION', CodeVar, ValVar) then
    begin
      if not RechercheVariable('G', 'GENERAL', CodeVar,   ValVar) then
      begin
        if not RechercheVariable('D', 'DOCUMENT', CodeVar,ValVar) then
        begin
          if not RechercheVariable('L', 'LIGNE', CodeVar,ValVar) then
          begin
            if not RechercheVariable('B', Article, CodeVar, ValVar) then
          end;
        end;
      end;
    end;

    SetControlText('BVA_VALEUR', ValVar);

  end;
  *}

end ;

function Tom_BVARIABLES.RechercheVariable(TypeVAR, Article, CodeVar : String; Var ValVar : Variant) : String;
Var Req : string;
    Q   : TQuery;
Begin

  Result := '';

  Req := 'SELECT BVA_VALEUR FROM BVARIABLES ';
  Req := Req + 'WHERE (BVA_TYPE="'        + TypeVar + '") ';
  Req := Req + 'AND (BVA_ARTICLE="'       + Article + '") ';
  Req := Req + 'AND (BVA_CODEVARIABLE="'  + CodeVar + '")';

  Q := OpenSql (Req,true);

  if not Q.eof then Result := Q.FindField('BVA_VALEUR').AsString;

  ferme(Q)

end;

procedure TOM_BVARIABLES.OnArgument ( S: String ) ;
var Critere  : String;
    Indice   : Integer ;
    ChampMul : String;
    ValMul   : string ;
begin
  Inherited ;

  SetControlVisible ('LIBELLEARTICLE', True);

  GetObjects;

  GetEventObjects;

  Action := Self.Action;

  // Chargement du critére de sélection
  Critere:=uppercase(Trim(ReadTokenSt(S))) ;
  TypeVariable := 'B';

  if Critere = 'APPLICATION' then
  Begin
    SetControlCaption ('LIBELLEARTICLE','Variables Applications');
    CodeVariable.ElipsisButton := False;
    TypeVariable := 'A';
    exit;
  end
  Else if Critere = 'GENERALES' Then
  Begin
    SetControlCaption ('LIBELLEARTICLE','Variables Générales');
    CodeVariable.ElipsisButton := False;
    TypeVariable := 'G';
    exit;
  end
  Else if Critere = 'DOCUMENTS' then
  Begin
    SetControlCaption ('LIBELLEARTICLE','Variables Document');
    CodeVariable.ElipsisButton := False;
    TypeVariable := 'D';
  end
  Else if (Critere = 'LIGNES') Or (Critere = 'LIGNES SSD') then
  Begin
    CodeVariable.ElipsisButton := False;
    if Critere = 'LIGNES'     then
    begin
      SetControlCaption ('LIBELLEARTICLE','Variables Ligne Document');
      TypeVariable := 'L';
    end;
    if Critere = 'LIGNES SSD' then
    begin
      SetControlCaption ('LIBELLEARTICLE','Variables Ligne sous-Détail');
      TypeVariable := 'LS';
    end;
    Bvalider.visible := False
  end;

  indice:=pos('=',Critere);
  if indice <> 0 then
  begin
    ChampMul:=copy(Critere,1,indice-1);
    ValMul:=copy(Critere,indice+1,length(Critere));
    if champMul = 'LIBELLEARTICLE' then SetControlCaption ('LIBELLEARTICLE',Valmul);
  end;

  setControlVisible('BLast', false);
  setControlVisible('BFirst', false);
  setControlVisible('BNext', false);
  setControlVisible('BPrev', false);

  FListe.Row := 1;

end ;

procedure TOM_BVARIABLES.OnClose ;
begin
  Inherited ;

end ;

procedure TOM_BVARIABLES.OnCancelRecord ;
begin
  Inherited ;
end ;

Procedure TOM_BVARIABLES.GetObjects;
begin

    Bdelete       := Ttoolbarbutton97(GetControl('BDELETE'));
    BValider      := TtoolBarButton97(GetControl('BVALIDER'));
    BInsert       := TToolBarButton97(GetControl('BINSERT'));
    BFerme        := TToolBarButton97(GetControl('BFERME'));
    //
    FListe        := THDBGrid(GetControl('FLISTE'));
    //
    CodeVariable  := THDBEdit(GetControl('BVA_CODEVARIABLE'));
    Libelle       := THDBEdit(GetControl('BVA_LIBELLE'));
    Valeur        := THDBEdit(GetControl('BVA_VALEUR'));

end;

Procedure TOM_BVARIABLES.GetEventObjects;
begin

  bdelete.OnClick   := DeleteOnClick;

  FListe.OnRowEnter := RowEnterOnclick;

end;

Procedure TOM_BVARIABLES.RowEnterOnclick(Sender: TObject);
begin

  Valeur.Text := Fliste.DataSource.DataSet.findfield('BVA_VALEUR').AsString;

end;

Procedure TOM_BVARIABLES.DeleteOnClick(Sender : Tobject);
Var StSQl   : string;
    OkDel   : Boolean;
    QQ      : TQuery;
begin

  OkDel := True;

  if PGIAsk('Confirmez-vous la suppression de cette variable', 'Suppression Variable')=MrNo then Exit;

  //Avant de supprimer une Variables il faut vérifier si la variable n'est pas sur un document...
  //Variable Document et Lignes...
  if (TypeVariable = 'D') or (TypeVariable='L') then
  begin
    STSQL := 'SELECT * FROM BVARDOC WHERE BVD_CODEVARIABLE="' + CodeVariable.text + '"';
    QQ := OpenSQL(StSQl, False, -1, '', True);
    if not QQ.EOf then
    begin
      OkDel := False;
      PGIInfo('Cette variable se trouve sur un document et ne peut-être supprimée', 'Suppression Variables ' + TypeVariable);
    end;
    Ferme(QQ);
  end;

  if OkDel then
  begin
    //
    StSQl := 'DELETE BVARIABLES WHERE BVA_TYPE="' + TypeVariable + '" AND BVA_CODEVARIABLE="' + CodeVariable.Text + '"';
    if ExecuteSQL(StSQL) < 0 then PGIInfo('Suite à un problème la suppression ne s''est pas effectuée', 'Suppression Variables');
    RefreshDB;
  end;

end;


Initialization
  registerclasses ( [ TOM_BVARIABLES ] ) ;
end.

