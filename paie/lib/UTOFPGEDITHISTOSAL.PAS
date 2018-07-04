{***********UNITE*************************************************
Auteur  ...... : FC
Créé le ...... : 18/04/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : UTOFPGEDITHISTOSAL ()
Mots clefs ... : TOF;UTOFPGEDITHISTOSAL
*****************************************************************
PT1   21/06/2007 FC V_72 FQ 14330 Accès à tort aux salariés confidentiels
}
Unit UTOFPGEDITHISTOSAL ;

Interface

uses
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls,
  {$IFDEF EAGLCLIENT}
  eQRS1, UtilEAgl,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} QRS1, EdtREtat,
  {$ENDIF}
  HCtrls, HEnt1, UTOF,
  HQry, UTob, P5Def, StrUtils, ParamSoc,EntPaie, PGOutils2;


Type
  TOF_UTOFPGEDITHISTOSAL = Class (TOF)
  private
    TobEtat, T : Tob;
    procedure OnClickSalarieSortie(Sender: TObject);
    Procedure PGRechSal (FF : TForm) ; 
    procedure ExitEdit(Sender: TObject);
  public
    procedure OnArgument (S : String ) ; override ;
    procedure OnUpdate; override;
    procedure OnClose; override ;
  end ;

Implementation


procedure TOF_UTOFPGEDITHISTOSAL.OnArgument (S : String ) ;
var
  Check : TCheckBox;
begin
  inherited;
  SetControlText('DOSSIER',GetParamSocSecur('SO_LIBELLE','000000'));

  Check := TCheckBox(GetControl('CKSORTIE'));
  if (Check <> nil) then
    Check.OnClick:=OnClickSalarieSortie;

  PGRechSal (Ecran);
end;

procedure TOF_UTOFPGEDITHISTOSAL.OnClose ;
begin
  Inherited ;
  FreeAndNil (TobEtat);
end ;

procedure TOF_UTOFPGEDITHISTOSAL.OnUpdate;
var
  Q: TQuery;
  Pages: TPageControl;
  LeWhere,Requete,StWhere : String;
  Longueur : integer;
  DateArret : TDateTime;
  StDateArret : String;
  {$IFDEF EAGLCLIENT}
  StPages : String;
  {$ENDIF}
  St : String;
begin
  Inherited ;
  FreeAndNil (TobEtat);

  Pages := TPageControl(GetControl('Pages'));

  TobEtat := Tob.Create('Leshistosalariés', nil, -1);

  LeWhere := RecupWhereCritere(TFQRS1(Ecran).Pages);

  Requete := 'SELECT PHD_PGINFOSMODIF,PHD_SALARIE,PHD_DATEAPPLIC,PHD_ANCVALEUR,PHD_NEWVALEUR,'
    + ' PHD_TYPEVALEUR,PHD_TABLETTE'
    + ' FROM PGHISTODETAIL,SALARIES'
    + ' WHERE PHD_SALARIE=PSA_SALARIE'
    + ' AND PHD_PGTYPEINFOLS="SAL"';
  if (LeWhere <> '') then
  begin
    Longueur := Length(LeWhere);
    Longueur := Longueur - 5;
    StWhere := MidStr(LeWhere,6,Longueur);
    Requete := Requete + ' AND ' + StWhere;
  end;

  //DEB PT1
  St := SQLConf('SALARIES');
  if St <> '' then
    Requete := Requete + ' AND ' + St;
  //FIN PT1

  // Gérer la case à cocher exclure
  if (GetControlText('CKSORTIE')='X') and (IsValidDate(GetControlText('DATEARRET'))) then
  Begin
    DateArret := StrtoDate(GetControlText('DATEARRET'));
    StDateArret := ' AND (PSA_DATESORTIE>="'+UsDateTime(DateArret)+'" OR PSA_DATESORTIE="'+UsdateTime(Idate1900)+'" OR PSA_DATESORTIE IS NULL) ';
    StDateArret := StDateArret + ' AND PSA_DATEENTREE <="'+UsDateTime(DateArret)+'"';
    Requete := Requete + StDateArret;
  End;

  Q := OpenSQL(Requete, True);
  while not Q.Eof do
  begin
    T := Tob.Create('Leshistosalariés', TobEtat, -1);
    T.AddChampSup('PHD_SALARIE', False);
    T.AddChampSup('PHD_PGINFOSMODIF', False);
    T.AddChampSup('PHD_DATEAPPLIC', False);
    T.AddChampSup('PHD_ANCVALEUR', False);
    T.AddChampSup('PHD_NEWVALEUR', False);
    T.AddChampSup('LIBHISTO', False);

    T.PutValue('PHD_SALARIE', Q.FindField('PHD_SALARIE').AsString);
    T.PutValue('PHD_PGINFOSMODIF', Q.FindField('PHD_PGINFOSMODIF').AsString);
    T.PutValue('PHD_DATEAPPLIC', Q.FindField('PHD_DATEAPPLIC').AsDateTime);
    T.PutValue('PHD_ANCVALEUR', Q.FindField('PHD_ANCVALEUR').AsString);
    T.PutValue('PHD_NEWVALEUR', Q.FindField('PHD_NEWVALEUR').AsString);
    if (Q.FindField('PHD_TYPEVALEUR').AsString = 'T') then
      T.PutValue('LIBHISTO',RechDom(Q.FindField('PHD_TABLETTE').AsString, Q.FindField('PHD_NEWVALEUR').AsString, False))
    else
      T.PutValue('LIBHISTO', '');
    Q.Next;
  end;
  Ferme(Q);

  TobEtat.Detail.Sort('PHD_SALARIE;PHD_PGINFOSMODIF;PHD_DATEAPPLIC');
  {$IFDEF EAGLCLIENT}
    StPages := AglGetCriteres(Pages, FALSE);
  {$ENDIF}
  TFQRS1(Ecran).LaTob:= TobEtat;
end;

procedure TOF_UTOFPGEDITHISTOSAL.OnClickSalarieSortie(Sender: TObject);
begin
  SetControlenabled('DATEARRET',(GetControltext('CKSORTIE')='X'));
  SetControlenabled('TDATEARRET',(GetControltext('CKSORTIE')='X'));
end;

procedure TOF_UTOFPGEDITHISTOSAL.PGRechSal(FF: TForm);
var i : Integer;
LeControl : TControl;
begin
  for i := 0 to FF.ComponentCount - 1 do
  begin
    LeControl := TControl(FF.Components[i]);
    if LeControl is THLabel then continue;
    if (LeControl is THEdit) AND ((pos('SALARIE', LeControl.Name) > 0) OR (pos('SALARIE_', LeControl.Name) > 0)) then
      begin
      THEdit(LeControl).OnExit := ExitEdit;
      end
      else Continue;
  end;
end;

procedure TOF_UTOFPGEDITHISTOSAL.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;

Initialization
  registerclasses ( [ TOF_UTOFPGEDITHISTOSAL ] ) ;
end.



