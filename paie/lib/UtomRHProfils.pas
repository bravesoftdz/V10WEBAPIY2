{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 19/01/2007
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : RHPROFILS
Mots clefs ... : TOM;RHPROFILS
*****************************************************************}
Unit UtomRHProfils ;

Interface

Uses Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
     eFiche,eFichList,
{$ENDIF}
     sysutils,HCtrls,HEnt1,UTOM,UTob ,Paramsoc;

Type
  TOM_RHPROFILS = Class (TOM)
    procedure OnUpdateRecord             ; override ;
    procedure OnNewRecord                ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField (F: TField)  ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnDeleteRecord             ; override ;
    end ;

Implementation

procedure TOM_RHPROFILS.OnUpdateRecord ;
begin
  Inherited ;
end;

procedure TOM_RHPROFILS.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_RHPROFILS.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_RHPROFILS.OnChangeField (F: TField);
begin
end;

procedure TOM_RHPROFILS.OnArgument ( S: String ) ;
var Q : TQuery;
    TobChamps,T : Tob;
    i : Integer;
    Libelle,Champ,Tablette : String;
begin
  Inherited ;
  Q := OpenSQL('SELECT PAI_LIBELLE,PAI_LIENASSOC,PAI_IDENT FROM PAIEPARIM WHERE PAI_LIBELLE<>""',True);
  TobChamps := Tob.Create('LesChamps',Nil,-1);
  TobChamps.LoadDetailDB('LesChamps','','',Q,False);
  Ferme(Q);
  For i := 1 to 5 do
  begin
    Champ := GetParamsoc('SO_PGRHPROFILCOMP'+IntToStr(i));
    If champ <> '' then T := TobChamps.FindFirst(['PAI_IDENT'],[StrToInt(Champ)],False)
    else T := Nil;
    If T <> Nil then
    begin
      Libelle := T.Getvalue('PAI_LIBELLE');
      Tablette := T.Getvalue('PAI_LIENASSOC');
      SetControlCaption('TPOF_NIVEAURHPRO'+IntToStr(i),Libelle);
      SetControlProperty('POF_NIVEAURHPRO'+IntToStr(i),'DataType',Tablette);
    end
    else
    begin
      SetControlVisible('TPOF_NIVEAURHPRO'+IntToStr(i),False);
      SetControlVisible('POF_NIVEAURHPRO'+IntToStr(i),False);
    end;
  end;
end ;

procedure TOM_RHPROFILS.OnDeleteRecord ;
begin
  Inherited ;
end ;


Initialization
  registerclasses ( [ TOM_RHPROFILS ] ) ;
end.

