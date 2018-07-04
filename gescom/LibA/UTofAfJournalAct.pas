unit UTofAfJournalAct;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
{$IFDEF EAGLCLIENT}
    Maineagl,eQRS1,
{$ELSE}
   {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} db, FE_Main,qrs1,
{$ENDIF}

      HCtrls,UTOF, TraducAffaire, Windows,Messages,
      Dicobtp,utofbaseetats,ParamSoc, ConfidentAffaire, EntGc;

Type
     TOF_JOURNALACT = Class (TOF_BASE_ETATS)
        procedure OnArgument(stArgument : String ) ; override ;
        procedure OnLoad ; override ;
        procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);override ;
        procedure ComboRupture1Change(Sender: TObject); override ;
        procedure ComboRupture2Change(Sender: TObject); override ;
        procedure ComboRupture3Change(Sender: TObject); override ;
        procedure ComboRupture4Change(Sender: TObject); override ;
        procedure ComboRupture5Change(Sender: TObject); override ;
        procedure TableauObjectsInvisibles(Crit:string; var iNbChamps:integer; var tbChamps:PString);override ;
        public
        EtatVisa    :   THEdit;
     END ;

const
NbTbChInv = 12 ;
TbChampsInvisibles : array[1..NbTbChInv] of string 	= (
          {1}        'ARS_LIBELLE',
          {2}        'TARS_LIBELLE',
          {3}        'ARS_LIBELLE_',
          {4}        'TARS_LIBELLE_',
          {5}        'ACT_FONCTIONRES',
          {6}        'TACT_FONCTIONRES',
          {7}        'ACT_FONCTIONRES_',
          {8}        'TACT_FONCTIONRES_',
          {9}        'ACT_TYPERESSOURCE',
          {10}       'TACT_TYPERESSOURCE',
          {11}       'COMPLEMENT',
          {12}       'STATISTIQUES'
          );


Procedure AFLanceFiche_JournalActivite;
Procedure AFLanceFiche_JournalClient;
Procedure AFLanceFiche_RefacturationAffaire;
Procedure AFLanceFiche_RefacturationRess;
Procedure AFLanceFiche_JournalActivite1Ress (Ressource, TypeArticle, DateDeb, DateFin : string);
Procedure AFLanceFiche_JournalClient1Tiers (Tiers, Affaire, TypeArticle, DateDeb, DateFin : string);

implementation



procedure TOF_JOURNALACT.OnArgument(stArgument : String );
  Var
  Critere, Champ, valeur, St  : String;
  xx : integer;
  bAvecPreinit : boolean;
Begin
bAvecPreinit := false;

Inherited;

// Gestion du cochage des champs RuptAff si la combo XX_RUPTURE = ACT_AFFAIRE
CocherSiValeurDansCombo(RuptAff1, 'ACT_AFFAIRE', ComboRupture1);
CocherSiValeurDansCombo(RuptAff2, 'ACT_AFFAIRE', ComboRupture2);
CocherSiValeurDansCombo(RuptAff3, 'ACT_AFFAIRE', ComboRupture3);
CocherSiValeurDansCombo(RuptAff4, 'ACT_AFFAIRE', ComboRupture4);
CocherSiValeurDansCombo(RuptAff5, 'ACT_AFFAIRE', ComboRupture5);

EtatVisa := THEdit(GetControl('ACT_ETATVISA'));
if (EtatVisa<>NIl) and (GetParamSoc('so_afVisaActivite') = False) then begin
   SetControlVisible('ACT_ETATVISA',False);
   SetControlVisible('TACT_ETATVISA',False);
   // Initialisation doublée dans le onload dans ce cas car un filtre par défaut pourrait écraser cette valeur
   SetControlText ('ACT_ETATVISA','ATT;VIS');
   end;
if (GetParamSoc('SO_AFTYPESAISIEACT') <> 'FOL') then // mcd 25/02/03
  begin
  SetControlVisible ('ACT_FOLIO',False);
  SetControlVisible ('TACT_FOLIO',False);
  SetControlVisible ('ACT_FOLIO_',False);
  SetControlVisible ('TACT_FOLIO_',False);
  end;

////////////////////
// PL le 13/06/03 : appel depuis la saisie d'activité, on passe une ressource ou un client/mission
// en paramètres
///////////////////
St := stArgument;
// Recup des critères
Critere := (Trim (ReadTokenSt (St)));
While (Critere <> '') do
    BEGIN
    if (Critere <> '') then
        BEGIN
        xx := pos ('=', Critere);
        if (xx <> 0) then
           begin
           Champ := copy (Critere, 1, xx - 1);
           Valeur := Copy (Critere, xx + 1, length (Critere) - xx);
           end;

        if (Ecran.Name = 'AFJOURNALACT') then
          begin
          if Champ = 'RESSOURCE' then
             begin
             SetControlText ('ACT_RESSOURCE', Valeur);
             SetControlText ('ACT_RESSOURCE_', Valeur);
             SetControlEnabled('ACT_RESSOURCE', false);
             SetControlEnabled('ACT_RESSOURCE_', false);
             SetControlEnabled('ARS_LIBELLE', false);
             SetControlEnabled('ARS_LIBELLE_', false);
             SetControlEnabled('ACT_FONCTIONRES', false);
             SetControlEnabled('ACT_FONCTIONRES_', false);
             TTabSheet(GetControl('COMPLEMENT')).TabVisible := false;
             end;
          end;

        if (Ecran.Name = 'AFJOURNALCLIENT') then
          begin
          if Champ = 'TIERS' then
             begin
             SetControlText ('ACT_TIERS', Valeur);
             SetControlText ('ACT_TIERS_', Valeur);
             SetControlEnabled('ACT_TIERS', false);
             SetControlEnabled('ACT_TIERS_', false);
             SetControlEnabled('T_LIBELLE', false);
             SetControlEnabled('T_LIBELLE_', false);
             end;
          if Champ = 'AFFAIRE' then
             begin
             SetControlText ('ACT_AFFAIRE', Valeur);
             SetControlText ('ACT_AFFAIRE_', Valeur);
//             CodeAffaireDecoupe (CodeAffaire, part0, Part1, Part2, Part3,Avenant,taConsult,False);
             SetControlEnabled('ACT_AFFAIRE', false);
             SetControlEnabled('ACT_AFFAIRE_', false);
             SetControlEnabled('AFF_ADMINISTRATIF', false);
             SetControlEnabled('AFF_RESPONSABLE', false);
             TTabSheet(GetControl('STATAFFAIRE')).TabVisible := false;
             end;
          end;

        if Champ = 'TYPEARTICLE' then
          begin
            THMultiValComboBox (GetControl ('ACT_TYPEARTICLE')).Text := Valeur;
            THMultiValComboBox (GetControl ('ACT_TYPEARTICLE')).Plus := 'AND (CO_CODE="' + Valeur + '")';
            SetControlEnabled('ACT_TYPEARTICLE', false);
            if (Valeur = 'FRA') or (Valeur = 'MAR') then
              begin
                TCheckBox (GetControl ('ACT_ACTIVITEEFFECT')).state := cbUnChecked;
                SetControlEnabled('ACT_ACTIVITEEFFECT', false);
                SetControlVisible('XX_VARIABLE3', false);
                SetControlVisible('TXX_VARIABLE3', false);
              end;
          end;

        if Champ = 'DATEDEB' then
          begin
             SetControlText ('ACT_DATEACTIVITE', Valeur);
          end;

        if Champ = 'DATEFIN' then
          begin
             SetControlText ('ACT_DATEACTIVITE_', Valeur);
          end;

        if Champ = 'PREINIT' then
          bAvecPreinit := true;
        END;

    Critere := (Trim (ReadTokenSt (St)));
    END;

  inherited;

if (Ecran.Name = 'AFJOURNALCLIENT') then
  begin
    if (GetControlText ('ACT_AFFAIRE1') <> '') then
      begin
        SetControlEnabled('ACT_AFFAIRE1', false);
        SetControlEnabled('ACT_AFFAIRE1_', false);
      end;
    if (GetControlText ('ACT_AFFAIRE2') <> '') then
      begin
        SetControlEnabled('ACT_AFFAIRE2', false);
        SetControlEnabled('ACT_AFFAIRE2_', false);
      end;
    if (GetControlText ('ACT_AFFAIRE3') <> '') then
      begin
        SetControlEnabled('ACT_AFFAIRE3', false);
        SetControlEnabled('ACT_AFFAIRE3_', false);
      end;
    if (GetControlText ('ACT_AVENANT') <> '') then
      begin
        SetControlEnabled('ACT_AVENANT', false);
        SetControlEnabled('ACT_AVENANT_', false);
      end;

  end;

if bAvecPreinit then
  TFQRS1(Ecran).loaded;
//////////////////////////// Fin PL le 13/06/03
setContrOlVisible ('BPROG',FALSE); // les état utilisent fct ConversionUnite.. pas OK plannification
End;

procedure TOF_JOURNALACT.ComboRupture1Change(Sender: TObject);
begin
Inherited;
CocherSiValeurDansCombo(RuptAff1, 'ACT_AFFAIRE', ComboRupture1);
end;

procedure TOF_JOURNALACT.ComboRupture2Change(Sender: TObject);
begin
Inherited;
CocherSiValeurDansCombo(RuptAff2, 'ACT_AFFAIRE', ComboRupture2);
end;

procedure TOF_JOURNALACT.ComboRupture3Change(Sender: TObject);
begin
Inherited;
CocherSiValeurDansCombo(RuptAff3, 'ACT_AFFAIRE', ComboRupture3);
end;

procedure TOF_JOURNALACT.ComboRupture4Change(Sender: TObject);
begin
Inherited;
CocherSiValeurDansCombo(RuptAff4, 'ACT_AFFAIRE', ComboRupture4);
end;

procedure TOF_JOURNALACT.ComboRupture5Change(Sender: TObject);
begin
Inherited;
CocherSiValeurDansCombo(RuptAff5, 'ACT_AFFAIRE', ComboRupture5);
end;

procedure TOF_JOURNALACT.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Inherited;
end;

procedure TOF_JOURNALACT.TableauObjectsInvisibles(Crit:string; var iNbChamps:integer; var tbChamps:PString);
begin
if (Crit='RESS') then
    begin
    iNbChamps := NbTbChInv;
    tbChamps := @TbChampsInvisibles;
    end;
end;

procedure TOF_JOURNALACT.OnLoad;
Var combo:ThValCOmboBox;
ii:integer;
begin
  inherited;
// mcd 31/07/02 pour prendre en compte l'obligationd e mettre '.X' dans le champ co_libre de
// afrupture  afin d'avoir une tablette sans zone fichier
//  (amis qui doit fonctionner si filtre (qui est fait sur cett valeur)==> rien ne doit être écrit dans le where
Combo :=  THValComboBox(GetControl('XX_VARIABLE2'));
If combo <> Nil then begin
   For ii:=0 to Combo.values.count-1 do begin
       If Combo.values[ii]='.'+IntToStr(ii) then Combo.values[ii]:=' ';
       end;
   end;



// PL le 11/12/01 : initialisé dans le onload en plus du onargument car si l'utilisateur enregistre un filtre par defaut sur un etat visa
// puis revient sur le même écran avec un autre etat visa dans le onargument, le filtre va écraser l'initialisation du onargument
if (EtatVisa<>NIl) then
if (EtatVisa.visible=false) then
   SetControlText ('ACT_ETATVISA','ATT;VIS');

end;


Procedure AFLanceFiche_JournalActivite;
begin
     // Si pas manager, alors restrictions
if (SaisieActiviteManager = FALSE) then
begin
    if VH_GC.RessourceUser = '' then
     begin
     PgiInfoAF ('Vous n''avez pas accès à cette commande.Le User courant n''a pas de droits ou pas de code associé','Impression Journal');
     exit;
     end;
end;
AGLLanceFiche ('AFF','AFJOURNALACT','','','');
end;

Procedure AFLanceFiche_JournalClient;
begin
AGLLanceFiche ('AFF','AFJOURNALCLIENT','','','');
end;

Procedure AFLanceFiche_RefacturationAffAire;
begin
AGLLanceFiche ('AFF','AFREFACTAFF','','','');
end;

Procedure AFLanceFiche_RefacturationRess;
begin
AGLLanceFiche ('AFF','AFREFACTRES','','','');
end;

Procedure AFLanceFiche_JournalActivite1Ress (Ressource, TypeArticle, DateDeb, DateFin : string);
  var
  arg : string;
begin
  arg := '';
  if (Ressource <> '') then
    arg := 'RESSOURCE=' + Ressource + ';';
  if (TypeArticle <> '') then
    arg := arg + 'TYPEARTICLE=' + TypeArticle + ';';
  if (DateDeb <> '') then
    arg := arg + 'DATEDEB=' + DateDeb + ';';
  if (DateFin <> '') then
    arg := arg + 'DATEFIN=' + DateFin + ';';

  arg := arg + 'PREINIT=;';

  AGLLanceFiche ('AFF','AFJOURNALACT','','', arg);
end;

Procedure AFLanceFiche_JournalClient1Tiers (Tiers, Affaire, TypeArticle, DateDeb, DateFin : string);
  var
  arg : string;
begin
  arg := '';
  if (Tiers <> '') then
    arg := 'TIERS=' + Tiers + ';';
  if (Affaire <> '') then
    arg := arg + 'AFFAIRE=' + Affaire + ';';
  if (TypeArticle <> '') then
    arg := arg + 'TYPEARTICLE=' + TypeArticle + ';';
  if (DateDeb <> '') then
    arg := arg + 'DATEDEB=' + DateDeb + ';';
  if (DateFin <> '') then
    arg := arg + 'DATEFIN=' + DateFin + ';';

  arg := arg + 'PREINIT=;';

  AGLLanceFiche ('AFF','AFJOURNALCLIENT','','', arg);
end;


Initialization
registerclasses([TOF_JOURNALACT]);
end.
