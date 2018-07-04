{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 04/06/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : TICKETINTEGR ()
Mots clefs ... : TOF;TICKETINTEGR
*****************************************************************}
Unit TICKETINTEGR_TOF ;

Interface

Uses
//unused     StdCtrls,
//unused     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
//unused     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     HDB,
//unused     Mul,
     FE_Main,
{$ELSE}
     emul,
     MaineAGL,
{$ENDIF}
     AGLInit,
     UTOB, //unused
//unused     forms,
     sysutils,
//unused     ComCtrls,
     HCtrls,
     HEnt1, 
     HMsgBox,
     PgOutils2,
     P5Def,
     Hqry,
//unused     UTOB,
     UTOF ;

Type
  TOF_TICKETINTEGR = Class (TOF)
    private
    WW, DateDeb,DateFin,SalarieD, SalarieF    : THEdit;
    Distribution, Etablissement           : THMultiValComboBox;
    procedure ActiveWhere(Sender: TObject);
    procedure GrilleDblClick(Sender: TObject);

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

procedure TOF_TICKETINTEGR.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_TICKETINTEGR.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_TICKETINTEGR.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_TICKETINTEGR.OnLoad ;
begin
  Inherited ;
    ActiveWhere (NIL);
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 11/06/2003
Modifié le ... : 11/06/2003
Description .. : Procédure OnArgument
Suite ........ : La période proposée est la période en cours, elle est
Suite ........ : modifiable.
Suite ........ : Sur l'onglet complément il est possible d'affiner la sélection 
Suite ........ : en utilisant les champs libres paramétrés au niveau des 
Suite ........ : paramètres société.
Suite ........ : 
Mots clefs ... : PAIE; TICKETINTEGR
*****************************************************************}
procedure TOF_TICKETINTEGR.OnArgument (S : String ) ;
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
  Distribution := THMultiValComboBox(GetControl('PRT_DISTRIBUTION'));
  Distribution.Value := '';
  Etablissement := THMultiValComboBox(GetControl('PRT_ETABLISSEMENT'));
  Etablissement.Value := '';
  SalarieD := THEdit(GetControl('PRT_SALARIE'));
  SalarieF := THEdit(GetControl('PRT_SALARIE_'));

  OkOk := RendPeriodeEnCours (ExerPerEncours,DebPer,FinPer);
  if OkOk then
  begin
   if  DateDeb <> NIL   then DateDeb.text:= DebPer;
   if  DateFin <> NIL   then DateFin.text:= FinPer;
  end;

  for Num := 1 to 4 do
  begin
    VisibiliteChampSalarie (IntToStr(Num),
                            GetControl ('PRT_TRAVAILN'+IntToStr(Num)),
                            GetControl ('TPRT_TRAVAILN'+IntToStr(Num)));
    VisibiliteChampLibreX(IntToStr(Num),GetControl ('PSA_LIBRE'+IntToStr(Num)),GetControl ('TPSA_LIBRE'+IntToStr(Num)));
    VisibiliteBoolLibreSal (IntToStr(Num),
                            GetControl ('PRT_BOOLLIBRE'+IntToStr(Num)));
  end;
  VisibiliteStat(GetControl ('PRT_CODESTAT'),GetControl ('TPRT_CODESTAT'));
end ;

procedure TOF_TICKETINTEGR.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_TICKETINTEGR.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_TICKETINTEGR.OnCancel () ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 11/06/2003
Modifié le ... :   /  /    
Description .. : procédure ActiveWhere
Suite ........ : Outre les critères de la fiche, la sélection des salariés se fera 
Suite ........ : sur la valeur des champs suivant :
Suite ........ : -le code distribution doit être renseigné  (fiche salarié 
Suite ........ : complémentaire, onglet tickets restaurant)
Suite ........ : -le type de ticket doit être renseigné (fiche salarié 
Suite ........ : complémentaire, onglet tickets restaurant)
Suite ........ : -le salarié ne doit pas être en suspension de paie
Suite ........ : -le salarié doit être présent dans la période (date entrée et 
Suite ........ : date sortie).
Mots clefs ... : PAIE; TICKETINTEGR
*****************************************************************}
procedure TOF_TICKETINTEGR.ActiveWhere(Sender: TObject);
begin
  WW.Text := 'PRT_DISTRIBUTION IS NOT NULL AND '+
             'PRT_DISTRIBUTION <> "" AND '+
             'PRT_TYPTICKET IS NOT NULL AND '+
             'PRT_TYPTICKET <> "" AND '+
             'PRT_DATEINTEG = "'+USDateTime(IDate1900)+'" AND '+
             'PRT_DATEDEBUT = "'+UsDateTime(StrToDate(DateDeb.Text))+'" AND '+
             'PRT_DATEFIN = "'+UsDateTime(StrToDate(DateFin.Text))+'"';
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 11/06/2003
Modifié le ... : 11/06/2003
Description .. : Procédure GrilleDblClick
Suite ........ : travail sur une TOB alimentée par les éléments de la liste 
Suite ........ : PGTICKETBULL (table CDETICKETS et SALARIES) 
Suite ........ : A partir de cette table la mise à jour de HISTOSAISRUB est 
Suite ........ : effectuée.
Suite ........ : La date d'intégration de la  table CDETICKETS est mise à 
Suite ........ : jour.
Mots clefs ... : PAIE; TICKETINTEGR
*****************************************************************}
procedure TOF_TICKETINTEGR.GrilleDblClick(Sender: TObject);
var
  St                                                    : String;
{$IFNDEF EAGLCLIENT}
    Liste:THDBGrid;
{$ELSE}
    Liste:THGrid;
{$ENDIF}

begin
{$IFDEF EAGLCLIENT}
  TheMulQ := TOB(Ecran.FindComponent('Q'));
  Liste:=THGrid(GetControl('FLISTE'));
{$ELSE}
  TheMulQ := THQuery(Ecran.FindComponent('Q'));
  Liste:=THDBGrid(GetControl('FLISTE'));
{$ENDIF}
{$IFDEF EAGLCLIENT}
  TFmul(Ecran).Q.TQ.Seek(Liste.Row-1) ;
  TFmul(Ecran).Q.First;
  if not TFmul(Ecran).Q.EOF then
{$ELSE}
  TheMulQ.First;
  if Not TheMulQ.EOF then
{$ENDIF}
  begin
    St := DateDeb.Text+';'+DateFin.Text+';I';
    {appel fiche de saisie pour modifier les valeurs qui seront intégrées}
    AGLLanceFiche ('PAY','SAITICKRESTAU',  '', '', St);
  end
  else
    PGIInfo('Aucune commande à intégrer','Intégration');
end;

Initialization
  registerclasses ( [ TOF_TICKETINTEGR ] ) ;
end.
