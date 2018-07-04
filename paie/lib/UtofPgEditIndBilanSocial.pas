unit UtofPgEditIndBilanSocial;

interface
uses
  Classes,
  Sysutils,
  ComCtrls,
  Utof,
  HQry,
  {$IFDEF EAGLCLIENT}
  eQRS1,
  {$ELSE}
  QRS1,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}    
  HCtrls,
  Utob;


type
  TOF_PGBSINDBILAN_ETAT = class(TOF)
    procedure OnArgument (St: string) ; override ;
    procedure OnLoad; override;
    procedure OnUpdate; override;
    procedure OnClose; override;
  private
    T: TOB ;
  end;

    TOF_PGBSINDBILAN_CUB = class(TOF)
    procedure OnArgument (St: string) ; override ;
    End;



implementation
uses p5def;


{ UTOFPGBSINDBILAN_ETAT }

procedure TOF_PGBSINDBILAN_ETAT.OnArgument(St: string);
Var
  Str: string;
  Q : TQuery;
  j : Integer;
begin
  inherited;
//  T := TOB.Create ('EtTaSoeur', Nil, -1) ;
  j:=0;
  Str := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE="PCL" ORDER BY CC_CODE';
  Q := OpenSql(str, True);
  While Not Q.eof do
    Begin
    Inc(j) ;
    SetControlText('CAT'+IntToStr(j),Q.FindField('CC_CODE').ASString);
    Q.Next;
    End;
  Ferme(Q);
end;

procedure TOF_PGBSINDBILAN_ETAT.OnClose;
begin
  inherited;
  FreeAndNil(T) ;
end;

procedure TOF_PGBSINDBILAN_ETAT.OnLoad;
begin
end;

procedure TOF_PGBSINDBILAN_ETAT.OnUpdate;
Var  Sql, St : String;
begin
  inherited;

  St  := RecupWhereCritere(TPageControl(GetControl('Pages')));

  Sql := 'SELECT SUM(PBC_VALCAT) PBC_VALCAT ,PBC_BSPRESENTATION,PBC_INDICATEURBS,'+
         'PBC_CATBILAN,PBC_DATEDEBUT,PBC_DATEFIN '+
         'FROM BILANSOCIAL '+St+
         'GROUP BY PBC_DATEDEBUT,PBC_DATEFIN,PBC_BSPRESENTATION,PBC_INDICATEURBS,PBC_CATBILAN '+
         'ORDER BY PBC_DATEDEBUT,PBC_DATEFIN,PBC_BSPRESENTATION,PBC_INDICATEURBS';

  if GetControlText('XX_RUPTURE1') <> '' then
    Sql := 'SELECT SUM(PBC_VALCAT) PBC_VALCAT ,PBC_BSPRESENTATION,PBC_INDICATEURBS,'+
           'PBC_CATBILAN,PBC_DATEDEBUT,PBC_DATEFIN,'+GetControlText('XX_RUPTURE1')+' '+
           'FROM BILANSOCIAL '+St+
           'GROUP BY PBC_DATEDEBUT,PBC_DATEFIN,PBC_BSPRESENTATION,'+GetControlText('XX_RUPTURE1')+',PBC_INDICATEURBS,PBC_CATBILAN '+
           'ORDER BY PBC_DATEDEBUT,PBC_DATEFIN,PBC_BSPRESENTATION,'+GetControlText('XX_RUPTURE1')+',PBC_INDICATEURBS';

  TFQRS1(Ecran).WhereSQL:=SQL;
end;

{ TOF_PGBSINDBILAN_CUB }

procedure TOF_PGBSINDBILAN_CUB.OnArgument(St: string);
var
Num : integer;
begin
  inherited;
for Num := 1 to 4 do
  begin
  VisibiliteChampSalarie(IntToStr(Num), GetControl('PBC_TRAVAILN' + IntToStr(Num)), GetControl('TPBC_TRAVAILN' + IntToStr(Num)));
  VisibiliteStat(GetControl('PBC_CODESTAT'), GetControl('TPBC_CODESTAT'));
  VisibiliteChampLibreSal(IntToStr(Num), GetControl('PBC_LIBREPCMB' + IntToStr(Num)), GetControl('TPBC_LIBREPCMB' + IntToStr(Num)))
  end;

end;

initialization
  registerclasses([TOF_PGBSINDBILAN_ETAT,TOF_PGBSINDBILAN_CUB]);
end.

//     'PBC_ETABLISSEMENT,PBC_INCREM,PBC_NODOSSIER,PBC_CODESTAT,PBC_COEFFICIENT,'+
//     'PBC_LIBELLEEMPLOI,PBC_QUALIFICATION,'+
//     'PBC_LIBREPCMB1,PBC_LIBREPCMB2,PBC_LIBREPCMB3,PBC_LIBREPCMB4,'+
//     'PBC_TRAVAILN1,PBC_TRAVAILN2,PBC_TRAVAILN3,PBC_TRAVAILN4 '+

