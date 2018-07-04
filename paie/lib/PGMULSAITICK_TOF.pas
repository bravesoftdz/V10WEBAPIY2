{***********UNITE*************************************************
Auteur  ...... : Paie - MF
Créé le ...... : 14/05/2003
Modifié le ... : 06/06/2003
Description .. : Source TOF de la FICHE : PGMULSAITICK ()
Suite ........ : Multi-critère de saisie de commande de tickets restaurant
Mots clefs ... : TOF;PGMULSAITICK
*****************************************************************}
{
 PT1 MF  25/02/2005 V_60 : la liste ne présente que les salariés pour lesquels
                           le code distribution affecté est existant.                                                        }
Unit PGMULSAITICK_TOF ;

Interface

Uses
//unused     StdCtrls, 
//unused     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     HDB,
     HQry, //unused
     FE_Main,
{$ELSE}
     MaineAGL,
     UTOB,
{$ENDIF}
//unused     forms,
     sysutils,
//unused     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     PgOutils2,
     P5Def,
//unused     HQry,
     AGLInit;

Type
  TOF_PGMULSAITICK = Class (TOF)
    private
    WW, DateDeb,DateFin,SalarieD, SalarieF    : THEdit;
    Distribution, Etablissement           : THMultiValComboBox;
    procedure GrilleDblClick (Sender: TObject);
    procedure ActiveWhere (Sender: TObject);
    public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Implementation

procedure TOF_PGMULSAITICK.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGMULSAITICK.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGMULSAITICK.OnUpdate ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie - MF
Créé le ...... : 06/06/2003
Modifié le ... :   /  /    
Description .. : Procédure OnLoad
Suite ........ : Lance l'ActiveWhere
Mots clefs ... : TOF;PGMULSAITICK
*****************************************************************}
procedure TOF_PGMULSAITICK.OnLoad ;
begin
  Inherited ;
  ActiveWhere (NIL);
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie - MF
Créé le ...... : 06/06/2003
Modifié le ... :   /  /    
Description .. : Procédure OnArgument
Mots clefs ... : TOF;PGMULSAITICK
*****************************************************************}
procedure TOF_PGMULSAITICK.OnArgument (S : String ) ;
var
{$IFDEF EAGLCLIENT}
  Grille                 : THGrid;
{$ELSE}
  Grille                 : THDBGrid;
{$ENDIF}
  DebPer,FinPer,ExerPerEncours              : String;
  OkOk                                      : Boolean;
  Num                                       : integer;

begin
  Inherited ;
{$IFDEF EAGLCLIENT}
  Grille:=THGrid (GetControl ('Fliste'));
{$ELSE}
  Grille:=THDBGrid (GetControl ('Fliste'));
{$ENDIF}
  if Grille <> NIL then Grille.OnDblClick := GrilleDblClick;

  DateDeb := ThEdit(getcontrol('XX_VARIABLEDEB'));
  DateFin := ThEdit(getcontrol('XX_VARIABLEFIN'));

  WW := THEdit (GetControl ('XX_WHERE'));
  Distribution := THMultiValComboBox(GetControl('PSE_DISTRIBUTION'));
  Distribution.Value := '';
  Etablissement := THMultiValComboBox(GetControl('PSA_ETABLISSEMENT'));
  Etablissement.Value := '';
  SalarieD := THEdit(GetControl('PSA_SALARIE'));
  SalarieF := THEdit(GetControl('PSA_SALARIE_'));

  OkOk := RendPeriodeEnCours (ExerPerEncours,DebPer,FinPer);
  if OkOk then
  begin
   if  DateDeb <> NIL   then DateDeb.text:= DebPer;
   if  DateFin <> NIL   then DateFin.text:= FinPer;
  end;

  for Num := 1 to 4 do
  begin
    VisibiliteChampSalarie (IntToStr(Num),
                            GetControl ('PSA_TRAVAILN'+IntToStr(Num)),
                            GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));
    VisibiliteChampLibreX(IntToStr(Num),GetControl ('PSA_LIBRE'+IntToStr(Num)),GetControl ('TPSA_LIBRE'+IntToStr(Num)));
    VisibiliteBoolLibreSal (IntToStr(Num),
                            GetControl ('PSA_BOOLLIBRE'+IntToStr(Num)));
  end;
  VisibiliteStat(GetControl ('PSA_CODESTAT'),GetControl ('TPSA_CODESTAT'));

end ;

procedure TOF_PGMULSAITICK.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PGMULSAITICK.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGMULSAITICK.OnCancel () ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie - MF
Créé le ...... : 06/06/2003
Modifié le ... :   /  /    
Description .. : procédure GrilleDblClick
Suite ........ : lance la fiche de saisie de la commande de tickets 
Suite ........ : restaurant (SAITICKRESTAU)
Mots clefs ... : TOF;PGMULSAITICK
*****************************************************************}
procedure TOF_PGMULSAITICK.GrilleDblClick(Sender: TObject);
var
  St                                            : String;
  Q                                             : TQuery;
begin

  St := 'select Distinct PRT_DATEDEBUT, PRT_DATEFIN FROM CDETICKETS WHERE (NOT '+
        '((PRT_DATEDEBUT < "'+UsDateTime(StrToDate(DateDeb.Text))+'" and '+
        ' PRT_DATEFIN < "'+UsDateTime(StrToDate(DateDeb.Text))+'") or '+
        '(PRT_DATEDEBUT > "'+UsDateTime(StrToDate(DateFin.Text))+'" and '+
        'PRT_DATEFIN > "'+UsDateTime(StrToDate(DateFin.Text))+'"))) AND'+
        '(PRT_DATEDEBUT <> "'+UsDateTime(StrToDate(DateDeb.Text))+'" OR '+
        'PRT_DATEFIN <>"'+UsDateTime(StrToDate(DateFin.Text))+'")';
  Q := OpenSql (St, TRUE);
  if not Q.EOF then
  begin
    PGIInfo('Il existe déjà des lignes de commande ','Saisie commande ');
    St := DateDeb.Text+';'+DateFin.Text+';C';
    AGLLanceFiche ('PAY','PGMULCDEXIST',  '', '', St);
  end;
  Ferme(Q);

{$IFDEF EAGLCLIENT}
  TheMulQ := TOB(Ecran.FindComponent('Q'));
{$ELSE}
  TheMulQ := THQuery(Ecran.FindComponent('Q'));
{$ENDIF}

  St := DateDeb.Text+';'+DateFin.Text+';S';
  AGLLanceFiche ('PAY','SAITICKRESTAU',  '', '', St);
  TheMulQ := NIL;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie - MF
Créé le ...... : 06/06/2003
Modifié le ... : 06/06/2003
Description .. : procédure ActiveWhere
Suite ........ : complément aux critères de sélection
Mots clefs ... : TOF;PGMULSAITICK
*****************************************************************}
procedure TOF_PGMULSAITICK.ActiveWhere(Sender: TObject);
// d PT1
var
  StDistribution                               : string;
  Q                                            : TQuery;
// f PT1
begin

// d PT1
  StDistribution := '';
  Q:= OpenSql('SELECT CC_ABREGE FROM CHOIXCOD '+
              'WHERE CC_TYPE="PCD"',True);
  While  not Q.eof do
  begin
    if (StDistribution = '') then
      StDistribution := ' AND (PSE_DISTRIBUTION = '
    else
      StDistribution := StDistribution + ' OR PSE_DISTRIBUTION = ' ;

    StDistribution := StDistribution+'"'+ Q.FindField('CC_ABREGE').AsString+'"';
    Q.Next;
  end;
  if (StDistribution <> '') then StDistribution := StDistribution + ')';
  ferme (Q);

// f PT1
  WW.Text := 'PSE_DISTRIBUTION IS NOT NULL AND '+
             'PSE_DISTRIBUTION <> "" AND '+
             'PSE_TYPTICKET IS NOT NULL AND '+
             'PSE_TYPTICKET <> "" AND '+
             '(PSA_SUSPENSIONPAIE <> "X") AND '+
             '(PSA_DATEENTREE <="'+UsDateTime(StrToDate(DateFin.Text))+'") AND '+
             '((PSA_DATESORTIE >="'+UsDateTime(StrToDate(DateDeb.Text))+'") OR '+
             '(PSA_DATESORTIE IS NULL) OR (PSA_DATESORTIE <= "'+UsDateTime(iDate1900)+'"))';
// d PT1
  if (StDistribution <> '') then
    WW.Text := WW.Text + StDistribution ;
// f PT1


end;
Initialization
  registerclasses ( [ TOF_PGMULSAITICK ] ) ;
end.
