{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 15/06/2001
Modifié le ... :   /  /
Description .. : Multi critère de gestion déportée des absences
Suite ........ : Accède aux absences salariés et aux absences que le
Suite ........ : responsable doit valider
Mots clefs ... : PAIE;ABSENCES;PGDEPORTEE
*****************************************************************
PT-1 SB 31/10/2001 V562 Rafraichissement de la liste après maj mvt en attente
PT-2 SB 31/10/2001 V562 Calcul du nb de mvt erronée retrait cond. AND PCN_SAISIEDEPORTEE="X"
                        Ajout variable Stperiode
PT-3 SB 02/01/2002 V571 Mvt à cheval sur 2 mois non visible
PT-4 SB 15/01/2002 V571 Affichage de la liste des salariés Responsable et collezerodevant
PT-5 SB 15/07/2002 V582 Critere de selection sur l'année en cours et non que le mois
PT-6 SB 15/07/2002 V582 RAZ du boolean de l'envoi de mail du responsable lors de la validation de l'abs salarié
PT-7 SB 15/07/2002 V582 Création d'un mvt abs du responsable ou administrateur pour le salarié
PT-8 SB 30/10/2002 V585 Recherche du nom prénom salarié à partir de la tablette
PT-9 SB 05/12/2002 V591 Ajout coche salariés sortis
PT-10 SB 05/12/2002 V591 Calcul des compteurs récap an cours aprés validation
PT-11-1 SB 21/07/2003 V_42 Affichage du mois de la date d'intégration
PT-11-2 SB 21/07/2003 V_42 Pour visualisation planning, récupération du matricule utilisateur
PT-12   SB 08/10/2003 V_42 Econges Spéc. CEGID Gestion des abences
PT-13   SB 12/11/2003 V_42 Econges Ajout paramètres
PT-14   SB 19/04/2005 V_60 FQ 12155 Econges Ajout contrôle
PT-15-1 SB 02/05/2005 V_60 FQ 12001 Modif objet PCN_VALIDRESP
PT-15-2 SB 02/05/2005 V_60 Ajout compteur enr. maj
PT-16   SB 05/07/2005 V_65 Ajout maj de la date de modif des mvts en validation
PT-17   SB 25/07/2005 V_65 FQ 12399 Ajout récapitulatif pour affichage assistante
PT-18   SB 26/07/2005 V_65 FQ 12415 Visibilité des zones libres pour l'assistante
PT-19   SB 25/07/2005 V_65 Affichage des sous niveaux hierarchiques
PT-20   SB 05/12/2005 V_65 FQ 12737 Ajout sélection du niveau hierarchique
PT-21   SB 19/06/2006 V_65 FQ 13231 Retrait des mvt absences annulées
PT-22   SB 17/10/2006 V_70 Refonte Fn pour utilisation portail => exporter vers pgcalendrier}
unit UtofPgEAbsenceMul;

interface
uses  Controls,Classes,sysutils, HTB97,stdctrls,
{$IFDEF EAGLCLIENT}
      UtileAGL,eMul,MaineAgl,HStatus,Utob,
{$ELSE}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,FE_Main,Mul,HStatus,
{$ENDIF}
      HCtrls,HQry,HEnt1,HMsgBox,UTOF,Paramsoc;

Type
     tof_PgEAbsenceMul = Class (TOF)
       private
       Ficprec,Stperiode,StSal       : string; //PT-9
       WW            : THEdit;
       QMul          : TQUERY;     // Query recuperee du mul
       vcbxMois, vcbxAnnee : THValComboBox;
       LaDated,Ladatef : TDateTime;
       GblNbUpdate     : integer;
       procedure ActiveWhere (Sender: TObject);
      public
       procedure OnArgument(Arguments : String ) ; override ;
       procedure OnLoad; override;
//       procedure Controledate (Sender: TObject);
       procedure btValidclick(sender : tobject);
       procedure BtInsertclick(sender : tobject);
       procedure btRecapclick(sender : tobject);
       procedure traiteliste(mode:string);
       procedure enableChampOngletModeAgl;
       procedure btPlanningclick(sender : tobject);
       procedure GrilleDblClick(Sender: TObject);
       procedure MajAbsenceExport (Mode : String);
       procedure OnExitSalarie(Sender: TObject);  //PT-4
       procedure ElipsisClickSal(Sender: TObject); //PT-4
       procedure ClickHierarchie(Sender: TObject);  { PT-20 }
     END ;

implementation
uses EntPaie,P5Util,P5Def,PGoutils2,PgOutilseAgl,PgPlanning,PgPlanningOutils,PgHierarchie,
      PgCalendrier;

{ La confection de la clause where depend du parametre et donc aussi de la fiche
  utilisée. Les listes sont différentes
  3 cas : Salarié        ==> absences du salariés
          Responsable    ==> absences des salariés qui ont comme responsable
          Administrateur ==> accès à toutes les absences
}
procedure tof_PgEAbsenceMul.ActiveWhere(Sender: TObject);
var Annee,Mois,JJ : WORD;
    ChbxH : TCheckBox;
    Hierar : THMultiValComboBox;
    StHierar : String;
begin
// initilalisation du code salarié en fonction de la connection
if WW=nil then exit;
WW.Text:='';
LaDated:=idate1900; LaDateF:=idate1900;

ChbxH := TCheckBox (GetControl ('CKHIERARCHIE'));   { PT-19 }
Hierar := THMultiValComboBox (GetControl ('HIERARCHIE'));  { PT-20 }
StHierar := '';
If Ficprec = 'SAL' then
   begin  // Cas absence du salarie
   if LeSalarie <> '' then WW.text := ' PCN_SALARIE = "'+LeSalarie+'" '
    else  // Anomalie donc ne rien afficher dans le Mul donc requete impossible
     WW.text := ' PCN_SALARIE = "0000000000" AND PCN_SALARIE <> "0000000000" ';
   end
   else
   if Ficprec = 'ASS' then   { DEB PT-17 }
     begin  // Cas absence du salarie    { DEB PT-19 }
     If Assigned(ChbxH) and (ChbxH.Checked) and (LeSalarie <> '') then
     Begin
     if Assigned(Hierar) Then StHierar := RendClauseHierarchie(Hierar.Text);  { PT-20 }
     WW.text := StHierar + 'AND (PSE_ASSISTABS="'+LeSalarie+'" '+
                'OR PSE_RESPONSABS IN (SELECT PGS_RESPONSABS FROM SERVICES,SERVICEORDRE '+
                'WHERE PGS_CODESERVICE=PSO_CODESERVICE '+
                'AND PSO_SERVICESUP IN (SELECT S.PGS_CODESERVICE FROM SERVICES S '+
                'WHERE S.PGS_SECRETAIREABS="'+LeSalarie+'"))) ';
     end
     else                               { FIN PT-19 }
     if LeSalarie <> '' then WW.text := ' PSE_ASSISTABS = "'+LeSalarie+'" '
     else  // Anomalie donc ne rien afficher dans le Mul donc requete impossible
     WW.text := ' PSE_ASSISTABS = "0000000000" AND PSE_ASSISTABS <> "0000000000" ';
     end
   else                      {FIN PT-17 }
   begin
   if Ficprec = 'RESP' then
      begin // Cas des Absences à valider pour le responsable  { DEB PT-19 }
      If Assigned(ChbxH) and (ChbxH.Checked) and (LeSalarie <> '') then
      Begin
      if Assigned(Hierar) Then StHierar := RendClauseHierarchie(Hierar.Text);   { PT-20 }
      WW.text := StHierar + 'AND (PSE_RESPONSABS="'+LeSalarie+'" '+
                 'OR PSE_RESPONSABS IN (SELECT PGS_RESPONSABS FROM SERVICES,SERVICEORDRE '+
                 'WHERE PGS_CODESERVICE=PSO_CODESERVICE '+
                 'AND PSO_SERVICESUP IN (SELECT S.PGS_CODESERVICE FROM SERVICES S '+
                 'WHERE S.PGS_RESPONSABS="'+LeSalarie+'"))) ';
      end
      else                                                     { FIN PT-19 }
      if LeSalarie <> '' then  WW.text := ' PSE_RESPONSABS = "'+LeSalarie+'"'
       else WW.text := ' PSE_RESPONSABS = "0000000000" AND PSE_RESPONSABS <> "0000000000"';
      end
      else // Cas administrateur doit tout voir sinon erreur ne voit rien
      if Ficprec <> 'ADM' then WW.text := ' PCN_SALARIE = "0000000000" AND PCN_SALARIE <> "0000000000" ';
   end;
StPeriode:=''; //PT-2
DecodeDate(Idate1900,Annee,Mois,JJ);
if vcbxMois.Value<>''  then Mois:=Trunc(StrToInt(vcbxMois.Value));
if vcbxAnnee.Value<>'' then Annee:=Trunc(StrToInt(vcbxAnnee.Text));
if ((vcbxMois.Value <>'') and ((VcbxMois.value <>'') and (vcbxAnnee.value <>''))) then // cas où mois renseigné alors année obligatoire
 begin
 LaDated:=EncodeDate(annee,mois,1);
 LaDatef:=EncodeDate(annee,mois,1);
 end
 else if vcbxAnnee.Value<>'' then
    begin
    LaDatef:=EncodeDate(annee,12,1);
    LaDated:=EncodeDate(annee,1,1);
    end;
if ((vcbxMois <> NIL) AND (WW <>NIL) AND (vcbxAnnee <> NIL)) then
  if Ladated > idate1900 then
     begin  //DEB PT-2
     StPeriode:=' AND (PCN_DATEFINABS <="'+UsDateTime (Findemois(Ladatef))+'"'+
         ' AND PCN_DATEFINABS >= "'+UsDateTime(Debutdemois(Ladated))+'")'; //PT-3 modif champ PCN_DATEDEBUTABS
     WW.Text:= WW.Text + StPeriode;
     end;  //FIN PT-2
//DEB PT-9
if ((Ficprec = 'ADM') AND (GetControlText('CKSORTIE')='-')) OR (Ficprec <> 'ADM') then
  Begin
  StSal:=' AND (PSA_DATESORTIE<="'+UsDateTime(idate1900)+'" '+
            'OR PSA_DATESORTIE is null '+
            'OR PSA_DATESORTIE>="'+UsDatetime(Vh_Paie.PGECabDateIntegration)+'")';
  WW.text:=WW.text+StSal;
  End
else
  StSal:='';
//FIN PT-9

end;

procedure tof_PgEAbsenceMul.OnArgument(Arguments: String);
var
Edit,FP      : thedit;
Bt           : ttoolbarbutton97;
Zone         : Tcontrol;
//Combo        : THValcombobox;
num          : integer;
Grille       : THGrid;
ChbxH        : TCheckBox;
begin
inherited ;
Ficprec := Arguments;
WW:=THEdit (GetControl ('XX_WHERE'));
bt := ttoolbarbutton97(getcontrol('TVALID'));
if bt <> nil then bt.onclick := btValidclick;
bt := ttoolbarbutton97(getcontrol('TRECAP'));
if bt <> nil then bt.onclick := btRecapclick;
bt := ttoolbarbutton97(getcontrol('TPLANNING'));
if bt <> nil then
  Begin
  if Ficprec <> 'SAL' then Bt.Visible:=TRue;
  bt.onclick := btPlanningclick;
  end;

if Ficprec <> 'SAL' then SetControlText('PCN_VALIDRESP','ATT');
{Zone := getcontrol('PCN_VALIDRESP'); PT-15-1 Mise en commentaire
Combo := THValcombobox(Zone);
if ((Combo <> nil) and (Combo.value = '') ) then Combo.Itemindex := 1; }

vcbxMois  := THValComboBox (GetControl ('CBXMOIS'));
vcbxAnnee := THValComboBox (GetControl ('CBXANNEE'));
Zone := getcontrol('CBXMOIS');
InitialiseCombo(Zone);
Zone := getcontrol('CBXANNEE');
InitialiseCombo(Zone);

FP := THedit(getcontrol('FICPREC'));
if FP <> nil then FP.text := Ficprec;


if Ficprec <> 'ASS' then { DEB PT-17 }
  Begin
  Grille:=THGrid(getcontrol('FLISTE'));
  if Grille <> nil then Grille.OnDblClick := GrilleDblClick;
  End;                   { FIN PT-17 }

bt := ttoolbarbutton97(getcontrol('BInsert'));
if bt <> nil then bt.onclick := btInsertclick;
{ DEB PT-13 }
if (FicPrec<>'SAL') AND (GetControl ('PSA_TRAVAILN1')<>nil) then  { PT-18 }
   Begin
   SetControlProperty('Pavance','TabVisible',True);
   SetControlProperty('PComplement','TabVisible',VH_PAIE.PGEcabHierarchie);
   SetControlVisible('TPSE_CODESERVICE',VH_PAIE.PGEcabHierarchie);
   SetControlVisible('PSE_CODESERVICE' ,VH_PAIE.PGEcabHierarchie);
   For Num := 1 to VH_Paie.PGNbreStatOrg do
     begin
     SetControlProperty('PComplement','TabVisible',True);
     if Num >4 then Break;
     VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)),GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));
     end;
   For Num := 1 to VH_Paie.PgNbCombo do
     begin
     SetControlProperty('TBChamplibre','TabVisible',True);
     if Num >4 then Break;
     VisibiliteChampLibreSal(IntToStr(Num),GetControl ('PSA_LIBREPCMB'+IntToStr(Num)),GetControl ('TPSA_LIBREPCMB'+IntToStr(Num)));
     end;
  End;
{ FIN PT-13 }

enableChampOngletModeAgl;
//DEB PT-4
Edit := THedit(getcontrol('PCN_SALARIE'));
if Edit<>nil then
  Begin Edit.OnExit :=OnExitSalarie;
  If FicPrec<>'ADM' then Edit.OnElipsisClick:=ElipsisClickSal;
  End; //FIN PT-4

SetControlvisible('CKSORTIE',(FicPrec='ADM')); //PT-9

{ DEB PT-20 }
if getparamsocsecur('SO_IFDEFCEGID',False)=True then
  Begin
  ChbxH := TCheckBox (GetControl ('CKHIERARCHIE'));           //PT-19
  if Assigned(ChbxH) then
    Begin
    ChbxH.Visible := VH_PAIE.PGEcabHierarchie;  //PT-19 { 13/10/2005 norme AGL }
    ChbxH.OnClick := ClickHierarchie;
    End;
  SetControlVisible('HIERARCHIE',VH_PAIE.PGEcabHierarchie);
  SetControlVisible('HHIERARCHIE',VH_PAIE.PGEcabHierarchie);
  End;
{ FIN PT-20 }

end;

procedure  tof_PgEAbsenceMul.btValidclick(sender : tobject);
begin
traiteListe('VAL');
end;

procedure  tof_PgEAbsenceMul.btPlanningclick(sender : tobject);
var
i : integer;
ListeSal,St,StFiche,StParam,StListe: string;
DD,FD : TDateTime ;
NiveauRupt : TNiveauRupture;
begin
ListeSal:=''; StFiche:='';  StParam:=''; StListe:='';
if (LeSalarie <> '') then
  Begin
  if Ficprec = 'SAL'  then
    StFiche := 'AND PSE_SALARIE = "'+LeSalarie+'" '
  else
    if Ficprec = 'RESP' then
      StFiche := 'AND PSE_RESPONSABS = "'+LeSalarie+'" ';
  End;

if GetControlText('PCN_SALARIE')<>'' then StParam:='OR PSE_SALARIE="'+GetControlText('PCN_SALARIE')+'" ';
if GetControlText('PSA_LIBELLE')<>'' then StParam:=StParam+'OR PSA_LIBELLE LIKE "'+GetControlText('PSA_LIBELLE')+'%" ';
If StParam<>'' then StParam:='AND ('+Copy(StParam,3,Length(StParam))+')';
St:=StParam;
for i:=0 to TFmul(Ecran).Fliste.NbSelected-1 do
  BEGIN
  TFmul(Ecran).Fliste.GotoLeBOOKMARK(i);
  {$IFDEF EAGLCLIENT}
  TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row-1) ;
  {$ENDIF}
  if Pos(TFmul(Ecran).Q.FindField('PCN_SALARIE').asstring,StListe)=0 then
  StListe:=StListe+'"'+TFmul(Ecran).Q.FindField('PCN_SALARIE').asstring+'",';
  END;
if StListe<>'' then
  St:='AND PSE_SALARIE IN ('+Copy(StListe,1,Length(StListe)-1)+')';
St:=StFiche+St;
//DEB PT-9
ListeSal:='WHERE (PSA_DATESORTIE<="'+UsDateTime(idate1900)+'" '+
              'OR PSA_DATESORTIE is null '+
              'OR PSA_DATESORTIE>="'+UsDatetime(Vh_Paie.PGECabDateIntegration)+'")';
if (St<>'') then ListeSal:=ListeSal+' AND '+Copy(St,4,Length(St)) ;
//FIN PT-9
//else if (St<>'') then  ListeSal:='WHERE ('+Copy(St,1,Length(St)-2)+')';
if (LaDated<=IDate1900) and  (VH_Paie.PGECabDateIntegration<>idate1900) then  //DEB PT-5
  Begin
  { DEB PT-11-1 }
  DD:=DebutDeMois(VH_Paie.PGECabDateIntegration);
  FD:=FinDeMois(VH_Paie.PGECabDateIntegration);
  { FIN PT-11-1 }
  End
else
  Begin
  DD:=LaDated;
  FD:=FindeMois(DD);
  End;                      //FIN PT-5
NiveauRupt.ChampsRupt[1]:=''; NiveauRupt.ChampsRupt[2]:='';
NiveauRupt.ChampsRupt[3]:=''; NiveauRupt.ChampsRupt[4]:='';
NiveauRupt.CondRupt:=''; NiveauRupt.NiveauRupt:=0;
PGPlanningAbsence(DD,FD,FicPrec,'','',ListeSal,LeSalarie,NiveauRupt); //PT5 //PT-11-2

end;
procedure  tof_PgEAbsenceMul.btRecapclick(sender : tobject);
begin
   AgllanceFiche ('PAY','MULRECAPSAL','','',Ficprec);
end;
// fonction de validation en bloc des absences
// On utilise la multi sélection, soit pour tout valider en bloc
// Vu avec JLD pour optimisation pour le pb de validation par lot
procedure tof_PgEAbsenceMul.traiteliste(mode:string);
VAR  I   : integer;
begin
{ DEB PT-14 }
If TFMul(Ecran).FListe.NbSelected < 1 then
   begin
   PGIBox ('Vous devez séléctionner les lignes d''absences en attente à valider.', Ecran.caption);
   exit;
   end;
{ FIN PT-14 }
{if GetControlText ('PCN_VALIDRESP') <> 'ATT' then  PT-15-2 Mise en commentaire
   begin
   PGIBox ('Vous ne pouvez valider que les mouvements en attente.', Ecran.caption);
   exit;
   end;  }
GblNbUpdate := 0 ;   
if PgiAsk('Confirmez-vous la validation des absences ?',Ecran.Caption)<>mrYes then exit; { PT-14 }
With TFMul(Ecran) do
   BEGIN
   if TFMul(Ecran).FListe.AllSelected then
      BEGIN
      InitMove(Q.RecordCount,'');
      Q.First;
      while Not Q.EOF do
        BEGIN
        MoveCur(False);
        MajAbsenceExport (Mode);
        Q.NEXT;
        END;
      TFMul(Ecran).FListe.AllSelected:=False;
      END ELSE
      BEGIN
      InitMove(TFMul(Ecran).FListe.NbSelected,'');
      for i:=0 to TFMul(Ecran).FListe.NbSelected-1 do
          BEGIN
          TFMul(Ecran).FListe.GotoLeBOOKMARK(i);
{$IFDEF EAGLCLIENT}
//          Q.TQ.Seek(FListe.Row-1) ;
          TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row-1) ;
{$ENDIF}
          MajAbsenceExport (Mode);
          MoveCur(False);
          END ;
      TFMul(Ecran).FListe.ClearSelected;
      END;
   FiniMove;
   if GblNbUpdate > 0  then PgiInfo(IntToStr(GblNbUpdate)+' mouvement(s) d''absence validé(s).Aucune anomalie n''a été détectée.',Ecran.Caption); { PT-15-2 }
   END ;
TFMUL(Ecran).BChercheClick (nil);    //PT-1
//PT-6 RAZ du boolean de l'envoi de mail du responsable lors de la validation de l'abs salarié
if (Ficprec='RESP') then
   PgMajIndicMailSalRAZ(LeSalarie);  { PT-22 }
//ExecuteSql('UPDATE DEPORTSAL SET PSE_EMAILENVOYE="-" WHERE PSE_SALARIE="'+LeSalarie+'"');
end;
// Fonction qui rend invisible les boutons en focntion du type de traitement
procedure tof_PgEAbsenceMul.enableChampOngletModeAgl;
begin
     SetcontrolEnabled('TREFUS'             , (Ficprec<>'ADM'));
//     SetControlVisible('TVALID'             , (Ficprec<>'ADM') AND (GetControlText ('PCN_VALIDRESP') = 'ATT'));
     SetControlEnabled('BOUVRIR'            , (Ficprec<>'ADM'));
end;

procedure tof_PgEAbsenceMul.OnLoad;
Var   Q                 : TQuery;
      LeNom{,LePrenom},st : String;
      Nbre              : Integer;
begin
inherited;
ActiveWhere (NIL);
{DEB PT-8 Mise en commentaire : remplacement par un rechdom }
LeNom:=RechDom('PGSALARIE',LeSalarie,False);
if FicPrec = 'SAL' then Ecran.Caption := 'Gestion de mes absences : '+LeNom{+' '+LePrenom}
 else if FicPrec = 'RESP' then Ecran.Caption := LeNom{+' '+LePrenom }+' : gestion des absences de mes collaborateurs'
  else if FicPrec = 'ADM' then Ecran.Caption := 'Supervision des absences déportées par l''administrateur';
{FIN PT-8}
// Test en fonction du type de traitement de l'affichage des controls
setcontrolVisible ('LSALARIE'      , (Ficprec<>'SAL'));  { PT-17 }
setcontrolVisible ('PCN_SALARIE'   , (Ficprec<>'SAL'));  { PT-17 }
setcontrolVisible ('LVALIDRESP'    , (Ficprec='RESP')or(Ficprec='ADM'));
//setcontrolVisible ('PCN_VALIDRESP' , (Ficprec='RESP')or(Ficprec='ADM')); PT-15-1
//setcontrolVisible ('TVALID'        , (Ficprec='RESP'));
setcontrolVisible ('TREFUS'        , (Ficprec='RESP'));


if Ficprec = 'RESP' then
   begin         //PT-2  AND PCN_SAISIEDEPORTEE="X"
   st := 'SELECT COUNT (*) NOMBRE FROM ABSENCESALARIE LEFT JOIN DEPORTSAL ON PSE_SALARIE=PCN_SALARIE '+
         'LEFT JOIN SALARIES ON PCN_SALARIE=PSA_SALARIE '+
         'WHERE ((PCN_TYPEMVT="ABS" AND PCN_SENSABS="-")OR PCN_TYPECONGE="PRI") AND PCN_VALIDRESP = "ATT"  '+
         'AND PCN_ETATPOSTPAIE <> "NAN" AND PSE_RESPONSABS="'+LeSalarie+'" '+StPeriode+StSal; //PT-9 { PT-21 }
   q := Opensql (St,true);
   if Not Q.EOF then //PORTAGECWAS
     Begin
     Nbre := Q.FindField ('NOMBRE').AsInteger;
     setcontrolvisible ('AVERT', True);
     setcontroltext    ('AVERT', TraduireMemoire ('Vous avez')+ ' '+ inttostr(Nbre)+' '+TraduireMemoire ('mouvement(s) à valider'));
     End;
   Ferme (Q);
   end;
UpdateCaption(TFmul(Ecran));
end;

// gestion du dblclick attention dans le cas de l'administrateur, ON PROPOSE une
// intervention identique au responsable
procedure tof_PgEAbsenceMul.GrilleDblClick(Sender: TObject);
var   Ordre,sal,etab,typemvt,lib,StAction,ValidResp,exportok : string;
begin
{$IFDEF EAGLCLIENT}
   QMUL:=TOB (TFMul(Ecran).Q.TQ);
{$ELSE}
   QMUL:=THQuery(Ecran.FindComponent('Q')) ;
{$ENDIF}
if QMUL.EOF then
begin
if FicPrec='SAL' then BtInsertclick (Sender);
exit;
end
else
begin
{$IFDEF EAGLCLIENT}
  TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row-1) ;   //Modif Eagl Sb 24/10/2001 : Repositionnement sur la grille
{$ENDIF}
if QMUL.Eof = true then begin BtInsertclick(nil); exit; end;
Ordre   := inttostr(QMUL.findfield('PCN_ORDRE').Asinteger);
Typemvt := QMUL.findfield('PCN_TYPEMVT').Asstring;
Sal     := QMUL.findfield('PCN_SALARIE').Asstring;
Lib     := QMUL.findfield('PSA_LIBELLE').Asstring;
Etab    := QMUL.findfield('PCN_ETABLISSEMENT').Asstring;
exportok:=QMUL.findfield('PCN_EXPORTOK').Asstring;
ValidResp:=QMUL.findfield('PCN_VALIDRESP').Asstring;
if Ficprec = 'SAL' then
  Begin
  if ValidResp='ATT' then StAction:='ACTION=MODIFICATION' else  StAction:='ACTION=CONSULTATION';
   AglLanceFiche ('PAY','EABSENCE',  'PCN_SALARIE='+sal, TYPEMVT+';'+sal+';'+Ordre,sal+';E;'+etab+';'+StAction+';;'+Ficprec)
  End
 else
  Begin
  if ExportOk='-' then  StAction:='ACTION=MODIFICATION' else  StAction:='ACTION=CONSULTATION';
  if IfMotifabsenceSaisissable(QMUL.findfield('PCN_TYPECONGE').Asstring,Ficprec) then StAction:='ACTION=MODIFICATION'; //PT12
  if (Ficprec = 'RESP') OR (Ficprec = 'ADM') then
    AglLanceFiche ('PAY','EABSENCE', 'PCN_SALARIE='+sal, TYPEMVT+';'+sal+';'+Ordre,sal+'!'+lib+';E;'+etab+';'+StAction+';;'+Ficprec);
  end;
end;
TFMUL(Ecran).BChercheClick (Sender);
end;
// gestion dE INSERT attention dans le cas de l'administrateur, il n'a pas le droit de créer
// des enregistrements, il le fera dans la paie mais dans la base de production de la paie
procedure tof_PgEAbsenceMul.BtInsertclick(sender: tobject);
var Sal,lib : String;
    OkOk : Boolean;
begin
//Lib     := TFMUL(Ecran).Q.findfield('PSA_LIBELLE').Asstring;
OkOk:=False;
if Ficprec = 'SAL' then
  Begin
  AglLanceFiche ('PAY','EABSENCE', 'PCN_SALARIE='+LeSalarie,LeSalarie,LeSalarie+';E;;ACTION=CREATION;;'+FICPREC);
  TFMUL(Ecran).BChercheClick (Sender);
  End
  else
  if (Ficprec = 'RESP') OR (Ficprec = 'ADM') then
  //DEB PT-7 Création d'un mvt abs du responsable ou administrateur pour le salarié
     if GetControlText('PCN_SALARIE')<>'' then
       Begin
       Sal:=GetControlText('PCN_SALARIE');
       lib:=RechDom('PGSALARIE',Sal,False);
       if (Ficprec = 'RESP') then  OkOk :=ExisteSql('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSABS="'+LeSalarie+'" AND PSE_SALARIE="'+Sal+'"');
       if (okok) OR (Ficprec = 'ADM') then
          Begin
          AglLanceFiche ('PAY','EABSENCE','PCN_SALARIE='+Sal ,Sal,Sal+'!'+lib+';E;;ACTION=CREATION;;'+FICPREC);
          TFMUL(Ecran).BChercheClick (Sender);
          End
       Else
          PgiBox('Le collaborateur saisi ne vous est pas affecté.',TFMUL(Ecran).caption);
       End
     else
       PgiBox('Veuiller renseigner le matricule du collaborateur.',TFMUL(Ecran).caption);
   //FIN PT-7
end;
// fonction de mise à jour du top absence validation du responsable
procedure tof_PgEAbsenceMul.MajAbsenceExport (Mode : String) ;
var st,Sal,TypMvt,ValidResp    : String;
    ordre            : Integer;
begin
ValidResp := TFmul(Ecran).Q.findfield('PCN_VALIDRESP').AsString;
if ValidResp <>'ATT' then exit; { PT-15-2 }
TypMvt := TFmul(Ecran).Q.findfield('PCN_TYPEMVT').asstring;
Sal    := TFmul(Ecran).Q.findfield('PCN_SALARIE').asstring;
Ordre  :=TFmul(Ecran).Q.findfield('PCN_ORDRE').asinteger;
GblNbUpdate := GblNbUpdate + PgMajAbsEtatValidSal(Sal,Mode,TypMvt,LeSalarie,ValidResp,Ordre);  { PT-22 }
(*st := 'UPDATE ABSENCESALARIE SET PCN_VALIDRESP = "'+Mode+'",PCN_VALIDABSENCE="'+LeSalarie+'", '+
      'PCN_DATEMODIF = "'+UsDateTime(Date)+'" '+ { PT-16 }
      'WHERE PCN_TYPEMVT = "'+TypMvt+'" AND PCN_SALARIE ="'+Sal+'" '+
      'AND PCN_ORDRE ='+IntToStr(Ordre);
GblNbUpdate := GblNbUpdate + ExecuteSql (st);   { PT-15-2 }
//PT-10 Recalcul des compteurs en cours lors de la validation
CalculRecapAbsEnCours(Sal);  *)
end;
//DEB PT-4
procedure tof_PgEAbsenceMul.OnExitSalarie(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;
procedure tof_PgEAbsenceMul.ElipsisClickSal(Sender: TObject);
begin
AfficheTabSalResp(Sender,LeSalarie);
end;
//FIN PT-4
{ DEB PT-20 }
procedure tof_PgEAbsenceMul.ClickHierarchie(Sender: TObject);
begin
SetcontrolEnabled('HIERARCHIE',(TCheckBox(Sender).Checked = True));
if TCheckBox(Sender).Checked = False then SetcontrolText('HIERARCHIE','');
end;
{ FIN PT-20 }

Initialization
registerclasses([tof_PgEAbsenceMul]);
end.
