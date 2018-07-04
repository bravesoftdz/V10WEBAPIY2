{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 26/08/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PASS_MONOMULTIDEP ()
Mots clefs ... : TOF;PASS_MONOMULTIDEP
*****************************************************************}
Unit PASS_MONOMULTIDEP_TOF ;

Interface
                  
Uses StdCtrls, 
     Controls, 
     Classes,
{$IFDEF EAGLCLIENT}
      MaineAGL,utileAGL,
{$ELSE}
     db,dbtables,Fiche,
{$ENDIF}
     AglInit,
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF,UTob,Vierge,EntGC,M3FP,UtilSoc ;

Type
  TOF_PASS_MONOMULTIDEP = Class (TOF)
    procedure LanceTraitement;
    procedure Annuler;
  end ;
                                         
const
  // libellés des messages
TexteMessage: array[1..6] of string 	= (
          {1}  'Recopie des établissements dans les dépôts de même code.'
          {2} ,'Il serait préférable de les supprimer.'
          {3} ,'Il serait préférable de le supprimer.'
          {4} ,'Recopie des titres des zones libres de l''établissement dans ceux des dépôts.'
          {5} ,'Recopie des tables Libres de l''établissement dans celles des dépôts.'
          {6} ,'Veuillez-vous reconnecter pour que les modifications soient prises en compte.'
              );
Implementation

procedure TOF_PASS_MONOMULTIDEP.LanceTraitement ;
var stInsertDepot,ListeDepot : string;
    NumEvt,i_ind : integer;
    EventsLog : TStringList;
    JNalCode : String;
    TOBJnal,TOBDepot : TOB ;
    QQ : TQuery;
begin
SetControlVisible('TRECAP',True);
SetControlVisible('LISTRECAP',True);
EventsLog := TStringList.Create;
EventsLog.Add('');
// suppression des dépôts ayant le même code qu'un établissement
ExecuteSQL('delete from depots where gde_depot in (select et_etablissement from etabliss)');
stInsertDepot:= 'insert into depots (gde_depot,gde_libelle,gde_abrege,gde_adresse1, gde_adresse2,gde_adresse3,gde_codepostal, gde_ville, gde_pays,'+
               'gde_telephone, gde_fax, gde_email, gde_blocnote, gde_utilisateur, gde_datecreation, gde_datemodif, gde_sursite, gde_sursitedistant,'+
               'gde_libredep1, gde_libredep2, gde_libredep3, gde_libredep4, gde_libredep5, gde_libredep6, gde_libredep7, gde_libredep8, gde_libredep9, '+
               'gde_libredepA, gde_datelibre1, gde_datelibre2, gde_datelibre3, gde_boollibre1, gde_boollibre2, gde_boollibre3, gde_charlibre1,'+
               'gde_charlibre2, gde_charlibre3, gde_vallibre1,gde_vallibre2, gde_vallibre3) select et_etablissement as gde_depot,et_libelle as  '+
               'gde_libelle,et_abrege as gde_abrege, et_adresse1 as gde_adresse1, et_adresse2 as gde_adresse2, et_adresse3 as '+
               'gde_adresse3,et_codepostal as gde_codepostal,et_ville as gde_ville,et_pays as gde_pays,et_telephone as gde_telephone,et_fax as '+
               'gde_fax,et_email as gde_email,et_blocnote as gde_blocnote, et_utilisateur as gde_utilisateur,et_datecreation as '+
               'gde_datecreation,et_datemodif as gde_datemodif,et_sursite as gde_sursite,et_sursitedistant as gde_sursitedistant,et_libreet1 as '+
               'gde_libredep1,et_libreet2 as gde_libredep2,et_libreet3 as gde_libredep3,et_libreet4 as gde_libredep4,et_libreet5 as '+
               'gde_libredep5,et_libreet6 as gde_libredep6,et_libreet7 as gde_libredep7,et_libreet8 as gde_libredep8,et_libreet9 as '+
               'gde_libredep9,et_libreetA as gde_libredepA,et_datelibre1 as gde_datelibre1,et_datelibre2 as gde_datelibre2,et_datelibre3 as '+
               'gde_datelibre3,et_boollibre1 as gde_boollibre1,et_boollibre2 as gde_boollibre2,et_boollibre3 as gde_boollibre3,et_charlibre1 as '+
               'gde_charlibre1,et_charlibre2 as gde_charlibre2,et_charlibre3 as gde_charlibre3,et_vallibre1 as gde_vallibre1,et_vallibre2 as '+
               'gde_vallibre2,et_vallibre3 as gde_vallibre3 '+
               'from etabliss';
ExecuteSQL(stInsertDepot);
// insertion d'un message dans le bloc note du journal des évenements
EventsLog.Add(TraduireMemoire(TexteMessage[1]));
// insertion du message dans la liste récapitulatif de la fenêtre
TListBox(GetControl('ListRecap')).Items.Add(TraduireMemoire(TexteMessage[1]));
// Si il y a des dépôts qui ne sont pas rattaché à un établissement, un message de conseillant de les supprimer
QQ := OpenSQL('Select * from depots where gde_depot not in (select et_etablissement from etabliss)',true);
if not QQ.EOF then
   begin
   TOBDepot := TOB.Create('liste depot',nil,-1);
   TOBDepot.LoadDetailDB('DEPOTS','','',QQ,False);
   for i_ind:=0 to TOBDepot.detail.count-1 do
      begin
      if ListeDepot<>'' then ListeDepot:=ListeDepot+',';
      ListeDepot:=ListeDepot+TOBDepot.detail[i_ind].GetValue('GDE_DEPOT');
      end;
   EventsLog.Add('');
   TListBox(GetControl('ListRecap')).Items.Add('');
   if TOBDepot.detail.count=1 then
      begin
      EventsLog.Add(TraduireMemoire('Le dépôt : ')+ListeDepot+TraduireMemoire(' ne correspond pas à un établissement'));
      TListBox(GetControl('ListRecap')).Items.Add(TraduireMemoire('Le dépôt : ')+ListeDepot+TraduireMemoire(' ne correspond pas à un établissement'));
      EventsLog.Add(TraduireMemoire(TexteMessage[3]));
      TListBox(GetControl('ListRecap')).Items.Add(TraduireMemoire(TexteMessage[3]));
      end
   else
      begin
      EventsLog.Add(TraduireMemoire('Les dépôts : ')+ListeDepot+TraduireMemoire(' ne correspondent pas à un établissement'));
      TListBox(GetControl('ListRecap')).Items.Add(TraduireMemoire('Les dépôts : ')+ListeDepot+TraduireMemoire(' ne correspondent pas à un établissement'));
      EventsLog.Add(TraduireMemoire(TexteMessage[2]));
      TListBox(GetControl('ListRecap')).Items.Add(TraduireMemoire(TexteMessage[2]));
      end;
   TOBDepot.free;
   end;
Ferme(QQ);

// suppression de toutes les lignes des titres des zones libres du dépôt et insertion à la place la valeur des titres des zones libres de l'établissement
ExecuteSQL('delete from CHOIXCOD where CC_TYPE="ZLI" and CC_CODE like "D%"');
ExecuteSQL('insert into CHOIXCOD (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) select "ZLI" as CC_TYPE,"D"+right(CC_CODE,2) as CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE from CHOIXCOD where CC_TYPE="ZLI" and CC_CODE like "E%"');
EventsLog.Add('');
EventsLog.Add(TraduireMemoire(TexteMessage[4]));
TListBox(GetControl('ListRecap')).Items.Add('');
TListBox(GetControl('ListRecap')).Items.Add(TraduireMemoire(TexteMessage[4]));
// suppression de toutes les lignes des tables libres du dépôt et insertion à la place la valeur des tables libres de l'établissement
ExecuteSQL('delete from CHOIXEXT where YX_TYPE like "D%"');
ExecuteSQL('insert into CHOIXEXT (YX_TYPE,YX_CODE,YX_LIBELLE,YX_ABREGE,YX_LIBRE) select "D"+right(YX_TYPE,2) as YX_TYPE,YX_CODE,YX_LIBELLE,YX_ABREGE,YX_LIBRE from CHOIXEXT where YX_TYPE like "E%"');
EventsLog.Add('');
EventsLog.Add(TraduireMemoire(TexteMessage[5]));
TListBox(GetControl('ListRecap')).Items.Add('');
TListBox(GetControl('ListRecap')).Items.Add(TraduireMemoire(TexteMessage[5]));
TListBox(GetControl('ListRecap')).Items.Add('');
TListBox(GetControl('ListRecap')).Items.Add(TraduireMemoire(TexteMessage[6]));
JNalCode:='OK';
// ??? EventsLog.Free;
// Journal des évènements --------

NumEvt := 0;
TOBJnal := TOB.Create('JNALEVENT', nil, -1);
TOBJnal.PutValue('GEV_TYPEEVENT', 'DEP');
TOBJnal.PutValue('GEV_LIBELLE', 'Gestion du passage de multi-dépôts en mono-dépôt');
TOBJnal.PutValue('GEV_DATEEVENT', Date);
TOBJnal.PutValue('GEV_UTILISATEUR', V_PGI.User);
QQ := OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT', true);
if not QQ.EOF then NumEvt := QQ.Fields[0].AsInteger;
Ferme(QQ);
inc(NumEvt);
TOBJnal.PutValue('GEV_NUMEVENT', NumEvt);
TOBJnal.PutValue('GEV_ETATEVENT', JNalCode);
TOBJnal.PutValue('GEV_BLOCNOTE', EventsLog.Text);
TOBJnal.InsertDB(nil);
TOBJnal.Free;
SetControlVisible('BTNFIN',True);
SetControlVisible('BTNANNULER',False);
SetControlVisible('BTNOK',False);
// Je passe cette variable à true car j'ai effectuer des modif
if (Ecran <> Nil) and (Ecran is TFVierge) then TFVierge(Ecran).Retour:='TRUE';
EventsLog.Free;
end ;

procedure TOF_PASS_MONOMULTIDEP.Annuler;
begin
// je passe cette variable à False car on annule l'opération, donc il faudra remettre en multi-dépôts
if (Ecran <> Nil) and (Ecran is TFVierge) then TFVierge(Ecran).Retour:='FALSE';
end;

procedure AGLLanceTraitement(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
     st1, st2 : string;
begin
F:=TForm(Longint(Parms[0])) ;
st1 := string(Parms[1]);
st2 := string(Parms[2]);
if (F is TFVierge) then MaTOF:=TFVierge(F).LaTOF else exit;
if (MaTOF is TOF_PASS_MONOMULTIDEP) then TOF_PASS_MONOMULTIDEP(MaTOF).LanceTraitement else exit;
end;

procedure AGLAnnuler(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     MaTOF  : TOF;
     st1, st2 : string;
begin
F:=TForm(Longint(Parms[0])) ;
st1 := string(Parms[1]);
st2 := string(Parms[2]);
if (F is TFVierge) then MaTOF:=TFVierge(F).LaTOF else exit;
if (MaTOF is TOF_PASS_MONOMULTIDEP) then TOF_PASS_MONOMULTIDEP(MaTOF).Annuler else exit;
end;

procedure initTOFPASS_MONOMULTIDEP();
begin
   RegisterAglProc('LanceTraitement',TRUE,1,AGLLanceTraitement);
   RegisterAglProc('Annuler',TRUE,1,AGLAnnuler);
end;

Initialization
   registerclasses ( [ TOF_PASS_MONOMULTIDEP ] ) ;
   initTOFPASS_MONOMULTIDEP();
end.
