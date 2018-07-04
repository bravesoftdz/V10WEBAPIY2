{***********UNITE*************************************************
Auteur  ...... : Guillaume Fontana
Créé le ...... : 01/09/2000
Modifié le ... : 06/09/2000
Description .. : TOM de la table HR_FAMRES (Famille de ressources)
Mots clefs ... :
*****************************************************************}
unit UtomBTFAMRES;
interface
uses
  SysUtils,
  Classes,
  Controls,
  StdCtrls,
  ComCtrls,
  Utob,
{$IFDEF EAGLCLIENT}              
  eFiche,
  MaineAGL,
{$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  db,
  Fiche,
  Fe_main,
{$ENDIF}
  HCtrls,
  UTOM,
  HENT1,
  HTB97,
  AGLINit,
  MsgUtil,
  M3FP;

type
  TOM_HRFAMRES = class(TOM)
    procedure AffichePeriodes;
    procedure OnArgument(S: string); override;
    procedure OnUpdateRecord; override;
    procedure OnNewRecord; override;
    procedure OnLoadRecord; override;
    procedure OnDeleteRecord; override;
    Procedure OnClose; override;

  private

    HFR_FAMRES				: THEdit;
    HFR_FAMRES1				: THEdit;
    HFR_LIBELLE				: THEdit;
    HFR_ABREGE				: THEdit;

    HTR_MODESURBOOK		: THValComboBox;
    HTR_ORDREPLANNING	: THSpinEdit;

    HFR_RYTHMEGESTION	: TComboBox;

    BPERIODES 				: TToolbarButton97 ;

    TobTypRes					: Tob;

    PremierPassage		: Boolean;

    procedure BPeriodes_OnClick(Sender: TObject);
    procedure HFR_RYTHMEGESTION_OnChange(Sender: TObject);
	  procedure HFR_FAMRES1_OnExit(Sender: TObject);
    procedure HTR_MODESURBOOK_OnChange(Sender: TObject);
		Procedure HTR_ORDREPLANNING_OnChange(Sender: TObject);
    function  ValidationPeriode(var ChampErreur: string): integer;
    procedure ValidationSousFamillle;

  end;

implementation

procedure TOM_HRFAMRES.OnArgument(S: string);
var Arg	: string;
begin
  inherited;

  PremierPassage := True;

  {Déclaration des zones de saisies pour gestion des evenements}
  HFR_RYTHMEGESTION := TComboBox(Ecran.FindComponent('HFR_RYTHMEGESTION'));
  HFR_RYTHMEGESTION.OnChange := HFR_RYTHMEGESTION_OnChange;

  //HTR_MODESURBOOK := THValComboBox(Ecran.FindComponent('HTR_MODESURBOOK'));
  //HTR_MODESURBOOK.OnChange := HTR_MODESURBOOK_OnChange;

  //HTR_ORDREPLANNING := THSpinEdit(Ecran.FindComponent('HTR_ORDREPLANNING'));
  //HTR_ORDREPLANNING.OnChange := HTR_ORDREPLANNING_OnChange;

  //BPERIODES  := TToolbarButton97(Ecran.FindComponent('BPERIODES'));
  //BPERIODES.OnClick := BPeriodes_OnClick;

  HFR_FAMRES := THEdit(Ecran.FindComponent('HFR_FAMRES'));
  HFR_FAMRES1:= THEdit(Ecran.FindComponent('HFR_FAMRES1'));
  HFR_FAMRES1.OnExit := HFR_FAMRES1_OnExit;

  HFR_LIBELLE:= THEdit(Ecran.FindComponent('HFR_LIBELLE'));
  HFR_ABREGE := THEdit(Ecran.FindComponent('HFR_ABREGE'));

  TobTypRes := Tob.Create('HRTYPRES' ,Nil, -1);
end;

procedure TOM_HRFAMRES.OnLoadRecord;
Var Req : String;
    QQ	: TQuery;
    Cod	: Integer;
begin
  inherited;

  HFR_FAMRES.Enabled := false;

  SetControlEnabled('HFR_FAMPRINCIPALE',True);

  //Chargement de la tob type de ressource associée
  Req := 'SELECT * FROM HRTYPRES WHERE HTR_RESSOURCE="' + GetField('HFR_FAMRES') + '"';
  Req := Req + ' AND HTR_NATURERES="TRE"';
  Req := Req + ' AND HTR_FAMRES="' + GetField('HFR_FAMRES') + '"';

  QQ := OpenSQL(req, TRUE);

  if Not QQ.Eof then
  	 Begin
     TobTypRes.selectDB('',QQ);
	   SetControlText('HTR_ORDREPLANNING',TobTypRes.GetValue('HTR_ORDREPLANNING'));
	   //SetControlText('HTR_NOMBRERES',TobTypRes.getvalue('HTR_NOMBRERES'));
	   SetControlText('HTR_MODESURBOOK',TobTypRes.getvalue('HTR_MODESURBOOK'));
	   end;

  Ferme(QQ);

  if GetField('HFR_FAMPRINCIPALE') = '-' then
  	 begin
     if ExisteSQL('Select HFR_FAMRES from HRFAMRES where HFR_FAMPRINCIPALE="X"') then
        begin
        SetControlEnabled('HFR_FAMPRINCIPALE',False);
        end;
     end;

  if (GetField('HFR_SURBOOKTYPRES') = '-') and (GetField('HFR_SURBOOKRESS') = '-') then
  	 begin
     SetControlEnabled('HFR_ALERTE1',False);
     SetControlEnabled('HFR_ALERTE2',False);
     SetControlEnabled('HFR_ALERTE3',False);
     end;

  if GetField('HFR_LIENEXTERIEUR') = '-' then
     begin
     SetControlEnabled('HFR_NLIENEXTERIEUR',False);
     end;

  AffichePeriodes;

  //Calcul du nombre de ressources
  Req := 'SELECT COUNT(ARS_RESSOURCE) As NBRES FROM RESSOURCE WHERE ARS_TYPERESSOURCE="' + GetField('HFR_FAMRES') + '"';

  QQ := OpenSQL(req, TRUE);

  if Not QQ.eof then
  	  Begin
     Cod := QQ.findField('NBRES').AsInteger;
	  SetControlText('HFR_NOMBRERES',IntToStr(cod));
     SetControlenabled('HFR_NOMBRERES', False)
	  end;

  ferme(QQ);

  Ecran.caption := 'Types de ressources';
  UpdateCaption(Ecran);

end;

procedure TOM_HRFAMRES.OnNewRecord;
begin

  inherited;
  // Si pas de gestion des liens exterieurs -> Nombre de liens exterieurs à 0
  if GetField('HFR_LIENEXTERIEUR') = '-' then
     SetField('HFR_NLIENEXTERIEUR',0);

  // Si pas de gestion du surbooking -> Niveaux d'alertes à 0
  if (GetField('HFR_SURBOOKTYPRES') = '-') and (GetField('HFR_SURBOOKRESS') = '-') then
     begin;
     SetField('HFR_ALERTE1',0);
     SetField('HFR_ALERTE2',0);
     SetField('HFR_ALERTE3',0);
  	 end;

  // Gestion Famille principale
  if ExisteSQL('Select HFR_FAMRES from HRFAMRES where HFR_FAMPRINCIPALE="X"') then
     begin
     SetField('HFR_FAMPRINCIPALE', '-');
     end;

end;

procedure TOM_HRFAMRES.OnUpdateRecord;
var
  ChampErreur: string;
  NumErreur: integer;
begin
  inherited;

  if GetField('HFR_RYTHMEGESTION') = '' then
     begin
     AfficheErreur(ecran.Name,'1','Types de ressources');
     SetFocusControl('HFR_RYTHMEGESTION');
     exit;
     end;

  // Si pas de gestion des liens exterieur -> Nombre de liens exterieurs à 0
  if GetField('HFR_LIENEXTERIEUR') = '-' then SetField('HFR_NLIENEXTERIEUR',0);

  // Si pas de gestion du surbooking -> Niveaux d'alertes à 0
  if (GetField('HFR_SURBOOKTYPRES') = '-') and (GetField('HFR_SURBOOKRESS') = '-') then
     begin
     SetField('HFR_ALERTE1',0);
     SetField('HFR_ALERTE2',0);
     SetField('HFR_ALERTE3',0);
     end;

  NumErreur := ValidationPeriode(ChampErreur);

  if NumErreur <> 0 then
     begin
		 AfficheErreur(ecran.Name,IntToStr(NumErreur),'Type de ressources');
     SetFocusControl(ChampErreur);
     Exit;
     end;

  ValidationSousFamillle;

end;

Procedure TOM_HRFAMRES.ValidationSousFamillle;
Begin

  TobTypRes.putvalue('HTR_RESSOURCE', GetControlText('HFR_FAMRES'));
  TobTypRes.putvalue('HTR_NATURERES', 'TRE');
  TobTypRes.putvalue('HTR_FAMRES', GetControlText('HFR_FAMRES'));
  TobTypRes.putvalue('HTR_TYPRES', '');
  TobTypRes.putvalue('HTR_LIBELLE', GetControlText('HFR_LIBELLE'));
  TobTypRes.putvalue('HTR_ABREGE', GetControlText('HFR_ABREGE'));
  //TobTypRes.putvalue('HTR_GESTIONRESS', GetControlText('HFR_GESTIONRESS'));
  TobTypRes.putvalue('HTR_SURBOOKRESS', GetControlText('HFR_SURBOOKRESS'));

  //mise en commentaire en attendant de connaitre la réelle fonction des zones
  {TobTypRes.putvalue('HTR_ALERTE1', GetControlText('HFR_ALERTE1'));
  TobTypRes.putvalue('HTR_ALERTE2', GetControlText('HFR_ALERTE2'));
  TobTypRes.putvalue('HTR_ALERTE3', GetControlText('HFR_ALERTE3'));
  TobTypRes.putvalue('HTR_ORDREPLANNING', GetControlText('HTR_ORDREPLANNING'));
  TobTypRes.putvalue('HTR_MODESURBOOK', GetControlText('HTR_MODESURBOOK'));
  TobTypRes.putvalue('HTR_NOMBRERES', StrToInt(GetControlText('HFR_NOMBRERES')));}
  //
  TobTypRes.putvalue('HTR_DATECREATION', Date);
  TobTypRes.putvalue('HTR_CREATEUR', V_PGI.User);
  TobTypRes.putvalue('HTR_DATEMODIF', date);
  TobTypRes.putvalue('HTR_UTILISATEUR', V_PGI.User);

  TobTypRes.setAllModifie(true);
  TobTypRes.InsertOrUpdateDB(true);

end;

procedure TOM_HRFAMRES.OnDeleteRecord;
var Req	: string;
begin
  inherited;

  //Chargement de la tob type de ressource associée
  Req := 'DELETE HRTYPRES WHERE HTR_RESSOURCE="' + GetField('HFR_FAMRES') + '"';
  Req := Req + ' AND HTR_NATURERES="TRE"';
  Req := Req + ' AND HTR_FAMRES="' + GetField('HFR_FAMRES') + '"';

  ExecuteSQL(Req);

  if ExisteSQL('Select HTR_FAMRES from HRTYPRES where HTR_FAMRES="' + GetField('HFR_FAMRES') + '"') then
     AfficheErreur(ecran.Name,'3','Types de ressources');

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Guillaume FONTANA
Créé le ...... : 18/04/2005
Modifié le ... : 18/04/2005
Description .. : Controle de la validité de la saisie des période
Mots clefs ... :
*****************************************************************}

function TOM_HRFAMRES.ValidationPeriode(var ChampErreur: string): integer;
begin

  Result := 0;
  ChampErreur := '';

  {Autres Modes de gestion , pas de gestion des périodes}
  SetField('HFR_MULTIPERIODE', '-');
  SetField('HFR_HRPERIODEDEBUT', '');
  SetField('HFR_HRPERIODEFIN', '');

  {Cas de gestion par période}
  if GetField('HFR_RYTHMEGESTION') <> '006' then exit;

  {Mono période : Période début = Période de fin }
  if GetField('HFR_MULTIPERIODE') = '-' then
     SetField('HFR_HRPERIODEFIN',GetField('HFR_HRPERIODEDEBUT'));

  {Mauvaises Saisies}
  if GetField('HFR_HRPERIODEDEBUT') = '' then
     Begin
     ChampErreur := 'HFR_HRPERIODEDEBUT';
     Result := 5;
     Exit;
  	 end;

	if GetField('HFR_HRPERIODEFIN') = '' then
     Begin
     ChampErreur := 'HFR_HRPERIODEFIN';
     Result := 6;
     exit;
     end;

  if (GetField('HFR_HRPERIODEFIN') < GetField('HFR_HRPERIODEDEBUT')) then
     begin
     ChampErreur := 'HFR_HRPERIODEDEBUT';
     Result := 7;
     end;

end;

procedure TOM_HRFAMRES.AffichePeriodes;
Var CodeGest : string;
begin

	CodeGest := GetControlText('HFR_RYTHMEGESTION');

  if Codegest = '006' then
	   begin
     SetControlVisible('HFR_MULTIPERIODE',True);
     SetControlVisible('HFR_HRPERIODEDEBUT',True);
     SetControlVisible('THFR_HRPERIODEDEBUT',True);
     SetControlVisible('HFR_HRPERIODEFIN',True);
     SetControlVisible('THFR_HRPERIODEFIN',True);
     end
  else
     begin
     SetControlVisible('HFR_MULTIPERIODE',False);
     SetControlVisible('HFR_HRPERIODEDEBUT',False);
     SetControlVisible('THFR_HRPERIODEDEBUT',False);
     SetControlVisible('HFR_HRPERIODEFIN',False);
     SetControlVisible('THFR_HRPERIODEFIN',False);
     end;

end;

procedure TOM_HRFAMRES.HFR_RYTHMEGESTION_OnChange(Sender: TObject);
begin
  AffichePeriodes;
end;

procedure TOM_HRFAMRES.BPeriodes_OnClick(Sender: TObject);
begin
  {On lance la fiche Periodes/Services}
  AGLLanceFiche('BTP', 'BTSERVICE', '', '', ActionToString(taModif));
  AvertirTable('BTSERVICES') ;
end;

procedure TOM_HRFAMRES.HFR_FAMRES1_OnExit(Sender: TObject);
var Lib : String;
		Abr	: String;
		Cod	: String;
begin

  Lib := '';
  Abr := '';
  Cod := '';

  Cod := HFR_FAMRES1.Text;

  if Cod = '' then exit;

	Lib := RechDom('AFTTYPERESSOURCE', Cod, False);
  Abr := RechDom('AFTTYPERESSOURCE', Cod, True);

  SetField('HFR_LIBELLE',Lib);
  SetField('HFR_ABREGE',Abr);

  HFR_RYTHMEGESTION.SetFocus;

end;

procedure TOM_HRFAMRES.OnClose;
begin
  inherited;

	TobTypRes.free;

end;

procedure TOM_HRFAMRES.HTR_MODESURBOOK_OnChange(Sender: TObject);
begin

  //if GetControlText('HFR_FAMRES') <> '' then ForceUpdate;

end;

procedure TOM_HRFAMRES.HTR_ORDREPLANNING_OnChange(Sender: TObject);
begin

  //if GetControlText('HFR_FAMRES') <> '' then ForceUpdate;

end;

initialization
  RegisterClasses([TOM_HRFAMRES]);
end.
