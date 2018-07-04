{***********UNITE*************************************************
Auteur  ...... : SB
Créé le ...... : 12/06/2001
Modifié le ... :   /  /
Description .. : Edition d'une DUE
Mots clefs ... : PAIE;DUE
*****************************************************************}
{PT 1 12/06/2001 SB Edition de la due
 PT 2 19/10/2001 SB Fiche de bug n°342
                    Mise en place de l'adresse de correspondance
 PT3 22/11/2001 SB Fiche de bug n°349
 PT3-1 Civilité de l'employeur
 PT3-2 Nationalité du salarié
 PT3-3 Medecine du travail
 PT3-4 Type de contrat de travail
 PT3-5 Telephone de la personne à contacter   *****Mis en commentaire le 26/05/2002 Fiche de bug n°349*****
 PT3-6 Temps de travail arrondi à l'entier supérieur
 PT4-1 18/02/2002 V571 SB Code medecine du travail erronée MED au lieu de MDT
 PT4-2 13/03/2002 V571 SB Fiche de Bug n°10026 Initialisation de champ à vide
 PT4-3 13/03/2002 V571 SB Fiche de Bug n°10026 Inversion nom marital nom de jeune fille
 PT5-1 16/05/2002 V582 JL Fiche de Bug n°349 Désactivation de CDD et date si CDI à l'initialisation
 PT5-2 16/05/2002 V582 JL Fiche de Bug n°349 Vérification du code pour tablette PGCODEQUALIF et PGCODEEMBAUCHE
 PT5-3 16/05/2002 V582 JL Fiche de bug n°349 Reprise de l'activité de l'établissement au lieu du libellé emploi
 PT6   08/08/2002 V585 SB FQ n°10155 Faute orthographe
 PT7   10/12/2002 V591 SB Nouveau PDF : zones supprimées
 PT8-1 10/12/2002 V591 SB FQ 10296 ajout champ PCI_SALAIREMOIS5
 PT8-2 10/12/2002 V591 SB FQ 10323 Bug contrat CDD et date prévisible d'embauche
 PT9   13/01/2003 V591 SB FQ 10356 Anomalie medecine du travail
 PT10  30/06/2003 V_42 PH Portage CWAS
 PT11  05/08/2003 V_42 SB FQ 10696 Si aucun contrat alors alimentation element salarié
 PT12  06/01/2004 V_42 SB FQ 10915 Ajout zone de saisie salaire brut
 PT13  22/04/2004 V_50 SB FQ 11058 Ajout zone de saisie Effectif, correction valeur par défaut
 PT14  23/04/2004 V_50 SB FQ 11265 Non reprise du code medecine du travail si égal à -1
 PT15  28/04/2004 V_42 SB FQ 10915 Ajout contrôle
 PT16  28/04/2004 V_50 SB FQ 10812 Intégration de la gestion des déclarants
 PT17  12/04/2006 V_65 JL modification clé annuaire
 PT18  07/09/2006 V_70 JL FQ 13396 Modif affichage médecine du travail
 PT19  08/01/2007 V_80 FCO FQ 13451 Pouvoir afficher les heures prévisibles d'embauche à 00:00 lorsqu'elles sont saisies (<> 01/01/1900)
 }

unit UTOFPGEditDue;

interface
uses StdCtrls,Controls,Classes,sysutils,ComCtrls,HTB97,
{$IFDEF EAGLCLIENT}
    UtileAgl,UTOB,
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     EdtREtat,
{$ENDIF}
     HCtrls,HEnt1,HMsgBox,HPdfviewer,ParamDat,UTof,LookUp;
Type
     TOF_PGDUE = Class (TOF)
       public
       procedure OnArgument(Arguments : String ) ; override ;
       private
       Matricule,Etab,Mode : string;
       DateEntree : TDateTime;
       procedure Imprimer(Sender: TObject);
       procedure OnChange(Sender: TObject);
       procedure GestionCheckOuiNon(CheckOui,CheckNon : TCheckBox);
       procedure InitSalarie(Arguments : String);
       procedure InitDue(Arguments : String);
       procedure DateElipsisclick(Sender: TObject);
       procedure InitTaille(ControlParent : TWinControl; taille : integer);
       procedure InitDateFait(Sender: TObject);
       procedure RensMedecine(Sender: TObject);
       procedure CPCorresp(Sender: TObject);  //PT 2
       procedure GestionCocheEmployeur(Sender: TObject);//PT3-1
      // procedure ChangeCodeMed(Sender: TObject);//PT3-3
       procedure GestionTypeContrat(Sender: TObject);//PT3-4
       procedure ControleValTablette(Sender:TObject); //PT5-2
       procedure ChangeEffAvEmb ( Sender : TOBJect );
       procedure ChangeSalaireBrut ( Sender : TOBJect ); //PT12
       procedure ChangeEff ( Sender : TOBJect );   { PT13 }
       procedure AffichDeclarant (Sender: TObject);  { PT16 }
       procedure AffichTelephone (Sender: TObject);  { PT16 }
     END ;
implementation

uses PgEditOutils,EntPaie; 

{ TOF_PGDUE }


procedure TOF_PGDUE.OnArgument(Arguments: String);
var
imprim                    : TToolBarButton97;
PDFDue                    :TPDFPreview;
CheckOui,CheckNon,Ck      : TCheckBox;
Edit                      : THEdit;
St                        : String;
begin
Matricule := ReadTokenSt(Arguments);
Mode:=Arguments;

{
 //Evenements
PDFDue :=TPDFPreview(GetControl('PDFDUE'));
if PDFDue <>nil then  Begin InitTaille(PDFDue,20);
   PDFDue.PDFPath:= VH_Paie.PGCheminRech+'\DUE.pdf'; End;
}

imprim := TToolBarButton97(GetControl('BImprimer'));
if imprim<>nil then imprim.OnClick:=Imprimer;

Edit:=THEdit(GetControl('DATEFAITEDIT'));
if Edit<>nil then Begin Edit.text:=DateToStr(Date);Edit.OnExit:=InitDateFait; Edit.OnDblClick:=DateElipsisclick; End;

Edit:=THEdit(GetControl('LIBELLEMED'));//PT18
if Edit<>nil then Edit.OnExit:=RensMedecine;

SetControlText('DATEFAIT',FormatCase(DateToStr(Date),'DATE',5));
//gestion des check Oui Non
{PT7 Mise en commentaire
CheckOui:=TCheckBox(GetControl('EXOOUI'));
CheckNon:=TCheckBox(GetControl('EXONON'));
If (CheckOui<>nil) and (CheckNon<>nil) then
   begin
   CheckOui.OnClick:=Onchange;
   CheckNon.OnClick:=Onchange;
   end;}
CheckOui:=TCheckBox(GetControl('ABBATOUI'));
CheckNon:=TCheckBox(GetControl('ABBATNON'));
If (CheckOui<>nil) and (CheckNon<>nil) then
   begin
   CheckOui.OnClick:=Onchange;
   CheckNon.OnClick:=Onchange;
   end;
CheckOui:=TCheckBox(GetControl('AGENTOUI'));
CheckNon:=TCheckBox(GetControl('AGENTNON'));
If (CheckOui<>nil) and (CheckNon<>nil) then
   begin
   CheckOui.OnClick:=Onchange;
   CheckNon.OnClick:=Onchange;
   end;
CheckOui:=TCheckBox(GetControl('LICENOUI'));
CheckNon:=TCheckBox(GetControl('LICENNON'));
If (CheckOui<>nil) and (CheckNon<>nil) then
   begin
   CheckOui.OnClick:=Onchange;
   CheckNon.OnClick:=Onchange;
   end;
CheckOui:=TCheckBox(GetControl('CBPREMIEROUI'));
CheckNon:=TCheckBox(GetControl('CBPREMIERNON'));
If (CheckOui<>nil) and (CheckNon<>nil) then
   begin
   CheckOui.OnClick:=Onchange;
   CheckNon.OnClick:=Onchange;
   end;
{PT7 mise en commentaire
CheckOui:=TCheckBox(GetControl('SALFR'));
CheckNon:=TCheckBox(GetControl('SAlEURO'));
If (CheckOui<>nil) and (CheckNon<>nil) then
   begin
   CheckOui.OnClick:=Onchange;
   CheckNon.OnClick:=Onchange;
   end;}


Edit := THEdit(GetControl('DATEFINCONTRAT'));
if Edit<>nil then edit.OnDblClick:=DateElipsisclick;

if Matricule<>'' then InitSalarie(Matricule);
if Matricule<>'' then InitDue(Matricule);
//Edit := THEdit(GetControl('CONTACT'));
//If Edit<>nil then Edit.OnChange:=ChangeContact; //PT3-5
Edit := THEdit(GetControl('CPCORRESP1'));//DEB PT 2
If Edit<>nil then Edit.OnChange:=CPCorresp;//FIN PT 2
// Deb PT3-1
if (GetControlText('CKETM')='X') then
   Begin
   SetControlEnabled('CKETM',True);
   SetControlEnabled('CKETMME',False);
   SetControlEnabled('CKETMLLE',False);
   End;
if (GetControlText('CKETMME')='X') then
   Begin
   SetControlEnabled('CKETM',False);
   SetControlEnabled('CKETMME',True);
   SetControlEnabled('CKETMLLE',False);
   End;
if (GetControlText('CKETMLLE')='X') then
   Begin
   SetControlEnabled('CKETM',False);
   SetControlEnabled('CKETMME',False);
   SetControlEnabled('CKETMLLE',True);
   End;
CK:=TCheckBox(GetControl('CKETM'));
If (Ck<>nil) then Ck.OnClick:=GestionCocheEmployeur;
Ck:=TCheckBox(GetControl('CKETMME'));
If (Ck<>nil) then Ck.OnClick:=GestionCocheEmployeur;
Ck:=TCheckBox(GetControl('CKETMLLE'));
If (Ck<>nil) then Ck.OnClick:=GestionCocheEmployeur;

//Edit := THEdit(GetControl('LBCODEMED'));   PT4-2 inutile
//If Edit<>nil then Edit.OnChange:=ChangeCodeMed; PT4-2
//Fin PT3-1
CK:=TCheckBox(GetControl('CBDURIND'));
If (Ck<>nil) then Ck.OnClick:=GestionTypeContrat;
CK:=TCheckBox(GetControl('CBDURDET'));
If (Ck<>nil) then Ck.OnClick:=GestionTypeContrat;
Edit:=THEdit(GetControl('QUALIF'));            //PT5-2
If Edit<>Nil Then Edit.OnExit:=ControleValTablette;
Edit:=THEdit(GetControl('EMBAUCHE'));           //PT5-2
If Edit<>Nil Then Edit.OnExit:=ControleValTablette;
Edit:=THEdit(GetControl('EFFECTIFAVEMB'));
If Edit<>Nil Then Edit.OnChange:=ChangeEffAvEmb;
Edit:=THEdit(GetControl('SALAIREBRUT'));   //PT12
If Edit<>Nil Then Edit.OnChange:=ChangeSalaireBrut;
Edit:=THEdit(GetControl('EFFECTIFSAIS')); { PT13 }
If Edit<>Nil Then Edit.OnChange:=ChangeEff;

{ DEB PT16 }
St := '(PDA_ETABLISSEMENT = "" OR PDA_ETABLISSEMENT LIKE "%'+Etab+'%") '+
      ' AND (PDA_TYPEATTEST = "" OR  PDA_TYPEATTEST LIKE "%DUE%" )  ' ;
SetControlProperty('DECLARANT', 'Plus' ,St );

Edit:=THEdit(GetControl('DECLARANT'));
if Edit<>nil then Edit.OnExit := AffichDeclarant;
Edit:=THEdit(GetControl('EDITTELEPHONE'));
if Edit<>nil then Edit.OnExit := AffichTelephone;
{ FIN PT16 }

end;


procedure TOF_PGDUE.GestionCheckOuiNon(CheckOui,CheckNon : TCheckBox);
begin
if (CheckOui<>nil) and (CheckNon<>nil) then
  begin
  If CheckOui.Checked=True then Begin CheckNon.Checked:=False; CheckNon.Enabled:=False  End
  else  CheckNon.Enabled:=True ;
  If CheckNon.Checked=True then Begin CheckOui.Checked:=False; CheckOui.Enabled:=False  End
  else  CheckOui.Enabled:=True ;
  End;
end;


procedure TOF_PGDUE.Imprimer(Sender: TObject);
Var Pages : TPageControl;
Edit : THEdit;
Lb : THLabel;
st {$IFDEF EAGLCLIENT} ,StPages {$ENDIF}  : string;
begin
Edit := THEdit(GetControl('QUALIF'));
Lb := THLabel(GetControl('THQUALIF'));
if (Edit<>nil) and (Lb<>nil) then  lb.Caption:=FormatCase(Edit.Text,'STR',5);
Edit := THEdit(GetControl('PSADURESSAI'));
Lb := THLabel(GetControl('THPSADURESSAI'));
if (Edit<>nil) and (Lb<>nil) then  lb.Caption:=FormatCase(Edit.Text,'STR',5);
Edit := THEdit(GetControl('PSADURHEB'));
Lb := THLabel(GetControl('THPSADURHEB'));
if (Edit<>nil) and (Lb<>nil) then  lb.Caption:=FormatCase(Edit.Text,'STR',5);
Edit := THEdit(GetControl('PSADURMENS'));
Lb := THLabel(GetControl('THPSADURMENS'));
if (Edit<>nil) and (Lb<>nil) then  lb.Caption:=FormatCase(Edit.Text,'STR',5);
Edit := THEdit(GetControl('PSADURAN'));
Lb := THLabel(GetControl('THPSADURAN'));
if (Edit<>nil) and (Lb<>nil) then  lb.Caption:=FormatCase(Edit.Text,'STR',5);
Edit := THEdit(GetControl('DATEFINCONTRAT'));
Lb := THLabel(GetControl('THDATEFINCONTRAT'));
if (Edit<>nil) and (Lb<>nil) then  lb.Caption:=FormatCase(Edit.Text,'DATE',5);
SetFocusControl('TBDUE'); { PT15 libère le focus pour execution des évènements sur zone }
{Bug 1 : Edition de la due
edition si due existante dans la table ATTESTATIONS
Modification de la requête par le source
Possibilité de rectifier sur l'état}
Edit := THEdit(GetControl('CODEMED'));
Lb := THLabel(GetControl('LBLCODEMED'));

if (Edit<>nil) and (Lb<>nil) then  lb.Caption:=FormatCase(Edit.Text,'STR',5);

st:='SELECT PSA_SALARIE FROM SALARIES WHERE PSA_SALARIE="'+Matricule+'" ';
Pages := TPageControl(GetControl('PAGES'));
//  PT10  30/06/2003 V_42 PH Portage CWAS
{$IFNDEF EAGLCLIENT}
LanceEtat('E','PAT','DUE',True,False,False,Pages,St,'',False);
{$ELSE}
StPages := AglGetCriteres (Pages, FALSE);
LanceEtat('E','PAT','DUE',True,False,False,NIL,St,'',False,0,StPages);
{$ENDIF}

end;

procedure TOF_PGDUE.InitDue(Arguments: String);
Var
QCount,QRech,QDue,QProf,QEntree,QSortie :TQuery;
PremSal : String;
DateDeb,DateFin{,DateEntree PT13 }: TDateTime;
nbJour :double;
Nbentree,NbSortie,MaxOrdre : integer;
begin
//  DateEntree:=Idate1900; PT13 Mise en commentaire => var. globale
  //AUTRES ELEMENTS SUR L'ENTREPRISE ET SUR L'EMPLOI initialisé à null
  SetControlText('EFFECTIFSAIS',''); { PT13 }
  SetControlText('EFFECTIF','');
  SetControlChecked('CBPREMIEROUI',False);
  SetControlChecked('CBPREMIERNON',False);
  SetControlText('EFFECTIFAV','');
  SetControlText('EFFECTIFAVEMB','');
  SetControlText('PSADURESSAI','');
  SetControlText('PSADURHEB','');
  SetControlText('PSADURMENS','');
  SetControlText('PSADURAN','');
  SetControlChecked('CBDURIND',False);
  SetControlChecked('CBDURDET',False);
  SetControlText('DATEFINCONTRAT','');
  {PT7 Mise en commentaire
  SetControlChecked('EXOOUI',False);
  SetControlChecked('EXONON',False);}
  SetControlText('SALAIRE','');
  SetControlText('SALAIREBRUT','0'); //PT12


  //AUTRES ELEMENTS SUR L'ENTREPRISE ET SUR L'EMPLOI
  QEntree:=OpenSql('SELECT COUNT(PSA_SALARIE) AS NBENTREE FROM SALARIES '+
                 'WHERE PSA_DATEENTREE<"'+USDateTime(DateEntree)+'" '+
                 'AND PSA_DATEENTREE<>"'+UsDateTime(iDate1900)+'" '+
                 'AND PSA_DATEENTREE IS NOT NULL AND PSA_ETABLISSEMENT="'+Etab+'"',True);
  if not QEntree.eof then NbEntree:=QEntree.FindField('NBENTREE').asinteger else NbEntree:=0;
  Ferme(QEntree);
  QSortie:=OpenSql('SELECT COUNT(PSA_SALARIE) AS NBSORTIE FROM SALARIES '+
                 'WHERE PSA_DATESORTIE<"'+USDateTime(DateEntree)+'" '+
                 'AND PSA_DATESORTIE<>"'+UsDateTime(iDate1900)+'" '+
                 'AND PSA_DATESORTIE IS NOT NULL AND PSA_ETABLISSEMENT="'+Etab+'"',True);
  if not QSortie.eof then NbSortie:=QSortie.FindField('NBSORTIE').asinteger else NbSortie:=0;
  Ferme(QSortie);
  if (NbEntree)-(NbSortie)>0 then
    SetControlText('EFFECTIFSAIS',FormatCase(IntToStr((NbEntree)-(NbSortie)),'STR',5)) { PT13 }
  else
    SetControlText('EFFECTIFSAIS',''); { PT13 }
  ChangeEff(nil);  { PT13 }

  QCount:=OpenSql('SELECT PSA_SALARIE,PSA_DATEENTREE FROM CONTRATTRAVAIL '+
                'LEFT JOIN SALARIES ON PCI_SALARIE=PSA_SALARIE '+
                'WHERE PCI_TYPECONTRAT="CDI" AND PSA_ETABLISSEMENT="'+Etab+'" '+
                'AND PSA_DATEENTREE<>"'+UsDateTime(iDate1900)+'" '+
                'AND PSA_DATEENTREE IS NOT NULL ORDER BY PSA_DATEENTREE',True);
  if not QCount.EOF Then //PORTAGECWAS
     PremSal:=QCount.FindField('PSA_SALARIE').AsString
  else
     PremSal:='';
  Ferme(QCount);
  if PremSal<>Matricule then
    begin
    SetControlChecked('CBPREMIEROUI',False);
    SetControlChecked('CBPREMIERNON',True);
    SetControlText('EFFECTIFAV','');
    SetControlText('EFFECTIFAVEMB','');
    end
  else
    begin
    SetControlChecked('CBPREMIEROUI',True);
    SetControlChecked('CBPREMIERNON',False);
    QEntree:=OpenSql('SELECT COUNT(PSA_SALARIE) AS NBENTREE FROM SALARIES '+
                   'WHERE PSA_DATEENTREE<"'+USDateTime(DateEntree)+'" '+
                   'AND PSA_DATEENTREE<>"'+UsDateTime(iDate1900)+'" '+
                   'AND PSA_DATEENTREE IS NOT NULL',True);
    if not QEntree.eof then NbEntree:=QEntree.FindField('NBENTREE').asinteger else NbEntree:=0;
    Ferme(QEntree);
    QSortie:=OpenSql('SELECT COUNT(PSA_SALARIE) AS NBSORTIE FROM SALARIES '+
                   'WHERE PSA_DATESORTIE<"'+USDateTime(DateEntree)+'" '+
                   'AND PSA_DATESORTIE<>"'+UsDateTime(iDate1900)+'" '+
                   'AND PSA_DATESORTIE IS NOT NULL',True);
    if not QSortie.eof then NbSortie:=QSortie.FindField('NBSORTIE').asinteger else NbSortie:=0;
    Ferme(QSortie);
    //SetControlText('EFFECTIFAV',FormatCase(IntToStr(NbEntree-NbSortie),'STR',5));
    SetControlText('EFFECTIFAVEMB',IntToStr(NbEntree-NbSortie));
    end;

  Qrech:=Opensql('SELECT MAX(PCI_ORDRE) AS MAXORDRE FROM CONTRATTRAVAIL WHERE PCI_SALARIE="'+matricule+'" ',True);
  if Not Qrech.eof then
    MaxOrdre:=QRech.FindField('MAXORDRE').asinteger
  else
    MaxOrdre:=0;
  Ferme(QRech);
  { DEB PT11 }
  if (MaxOrdre>0) then
    QDue:=OpenSql('SELECT PCI_TYPECONTRAT,PCI_DEBUTCONTRAT,PCI_FINCONTRAT,'+
                'PCI_ESSAIDEBUT,PCI_ESSAIFIN,PCI_RNVESSAIDEB,PCI_RNVESSAIFIN,'+
{PT8-1}         'PCI_SALAIREMOIS1+PCI_SALAIREMOIS2+PCI_SALAIREMOIS3+PCI_SALAIREMOIS4+PCI_SALAIREMOIS5 AS SAL,'+
                'PCI_HORAIREMOIS HORAIREMOIS,PCI_HORHEBDO HORHEBDO,PCI_HORANNUEL HORANNUEL FROM CONTRATTRAVAIL '+
                'WHERE PCI_SALARIE="'+Matricule+'" AND PCI_ORDRE='+IntToStr(MaxOrdre)+' ',True)
  else
    QDue:=OpenSql('SELECT PSA_HORAIREMOIS HORAIREMOIS,PSA_HORHEBDO HORHEBDO,PSA_HORANNUEL HORANNUEL, '+
                'PSA_SALAIREMOIS1+PSA_SALAIREMOIS2+PSA_SALAIREMOIS3+PSA_SALAIREMOIS4+PSA_SALAIREMOIS5 AS SAL '+
                'FROM SALARIES '+
                'WHERE PSA_SALARIE="'+Matricule+'" ',True);
  If Not QDue.EOF then //PORTAGECWAS
    Begin
    IF MaxOrdre > 0 then
       Begin
       DateDeb:=QDue.FindField('PCI_ESSAIDEBUT').AsDateTime;
       DateFin:=QDue.FindField('PCI_ESSAIFIN').AsDateTime;
       if (DateDeb<>idate1900) and (DateFin<>idate1900)  then
         begin
         nbJour:=DateFin-DateDeb;
         SetControlText('PSADURESSAI',FloatToStr(nbJour));
         end;
       //PT8-2 Affectation date début contrat à date embauche..
       if QDue.FindField('PCI_DEBUTCONTRAT').AsDateTime>idate1900 then
         SetControlText('PSADATEEMBAUCHE',FormatCase(DateToStr(QDue.FindField('PCI_DEBUTCONTRAT').AsDateTime),'Date',5));

       if QDue.FindField('PCI_TYPECONTRAT').AsString='CDI'then
          begin
          SetControlChecked('CBDURIND',True);
          SetControlEnabled('DATEFINCONTRAT',False);  // PT5-1
          SetControlEnabled('CBDURDET',False);
          end
       else
          Begin
          SetControlChecked('CBDURIND',False);
          if QDue.FindField('PCI_TYPECONTRAT').AsString='CCD'then //PT8-2 CDD au lieu de CCD
            begin
            SetControlChecked('CBDURDET',True);
            SetControlText('DATEFINCONTRAT',DateToStr(QDue.FindField('PCI_FINCONTRAT').AsDateTime));
            SetControlEnabled('CBDURIND',False);    // PT5-1
            end
          else
            begin
            SetControlChecked('CBDURDET',False);
            SetControlText('DATEFINCONTRAT','');
            end;
          End;
       End; //End MaxORdre
    SetControlText('PSADURHEB',FloatToStr(Arrondi(QDue.FindField('HORHEBDO').AsFloat,0)));    //PT3-6
    SetControlText('PSADURMENS',FloatToStr(Arrondi(QDue.FindField('HORAIREMOIS').AsFloat,0)));//PT3-6
    SetControlText('PSADURAN',FloatToStr(Arrondi(QDue.FindField('HORANNUEL').AsFloat,0)));    //PT3-6
    SetControlText('SALAIRE',FormatCase(QDue.FindField('SAL').AsString,'NBR',5));
    if IsNumeric(QDue.FindField('SAL').AsString) then //PT12
      SetControlText('SALAIREBRUT',QDue.FindField('SAL').AsString);
    END;
  Ferme(QDue);
  { FIN PT11 }

  QProf:=OpenSql('SELECT PSA_PROFILTPS,PPS_PROFIL FROM SALARIES '+
                 'LEFT JOIN PROFILSPECIAUX ON PPS_CODE=PSA_SALARIE AND PPS_ETABSALARIE="-" '+
                 'WHERE PSA_SALARIE="'+Matricule+'" ',True);
  if Not QProf.EOF then //PORTAGECWAS
    Begin
    {PT7 Mise en commentaire
    if QProf.FindField('PSA_PROFILTPS').asstring='017' then
       begin
       SetControlChecked('EXOOUI',True);
       SetControlChecked('EXONON',False);
       end
    else
       begin
       SetControlChecked('EXOOUI',False);
       SetControlChecked('EXONON',True);
       end;  }
    if QProf.FindField('PPS_PROFIL').asstring='028' then
       begin
       SetControlChecked('ABBATOUI',True);
       SetControlChecked('ABBATNON',False);
       end
    else
       begin
       SetControlChecked('ABBATOUI',False);
       SetControlChecked('ABBATNON',True);
       end;
    End;
  Ferme(Qprof);
  { DEB PT16 }
  QRech:=OpenSql('SELECT PDA_DECLARANTATTES FROM DECLARANTATTEST '+
                'WHERE (PDA_ETABLISSEMENT = "" OR PDA_ETABLISSEMENT LIKE "%'+Etab+'%") '+
                'AND (PDA_TYPEATTEST = "" OR  PDA_TYPEATTEST LIKE "%DUE%" )  '+
                'ORDER BY PDA_ETABLISSEMENT DESC',True);
  If Not QRech.eof then
    Begin
    SetControlText('DECLARANT' ,QRech.FindField('PDA_DECLARANTATTES').AsString);
    AffichDeclarant(nil);
    End;
  Ferme(QRech);
  { FIN PT16 }
end;

procedure TOF_PGDUE.InitSalarie(Arguments: String);
var
QSal : TQuery;
Heure,st : string;
DHeure : TDateTime;
Hour, Min, Sec, MSec: Word;
begin

QSal:=OpenSql('SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_ADRESSE1,PSA_ADRESSE2,'+
   'PSA_ADRESSE3,PSA_CODEPOSTAL,PSA_VILLE,PSA_LIBELLEEMPLOI,PSA_NUMEROSS,'+
   'PSA_NOMJF,PSA_PRENOMBIS,PSA_DATENAISSANCE,PSA_COMMUNENAISS,PSA_DEPTNAISSANCE,'+
   'PSA_PAYSNAISSANCE,PSA_NATIONALITE,PSA_SEXE,PSA_DATEENTREE,PSA_HEUREMBAUCHE,'+
   'PSA_CIVILITE,ET_ETABLISSEMENT,ET_FAX,'+
   'ET_LIBELLE,ET_ADRESSE1,ET_ADRESSE2,ET_ADRESSE3,ET_CODEPOSTAL,ET_VILLE,'+
   'ET_ACTIVITE,'+ //PT5-3
   'ET_TELEPHONE,ET_SIRET,ET_FAX,ET_APE,ETB_MEDTRAVGU, '+
   'ANN_NOM1,ANN_ALRUE1,ANN_ALRUE2,ANN_ALRUE3,ANN_ALVILLE,ANN_ALCP '+
   'FROM SALARIES '+
   'LEFT JOIN ETABLISS ON PSA_ETABLISSEMENT=ET_ETABLISSEMENT '+
   'LEFT JOIN ETABCOMPL ON PSA_ETABLISSEMENT=ETB_ETABLISSEMENT '+
   'LEFT JOIN ANNUAIRE ON ANN_GUIDPER=ETB_MEDTRAVGU AND ANN_TYPEPER="MED" '+ //PT4-1
   'WHERE PSA_SALARIE="'+Arguments+'" ',TRUE);

If Not QSal.EOF then //PORTAGECWAS
   BEGIN
   DateEntree:=QSal.FindField('PSA_DATEENTREE').AsDateTime;
   Etab :=QSal.FindField('ET_ETABLISSEMENT').AsString;
   SetControlText('ETLIBELLE',QSal.FindField('ET_LIBELLE').asstring);
   St:=QSal.FindField('ET_ADRESSE1').asstring+' '+QSal.FindField('ET_ADRESSE2').asstring+' '+QSal.FindField('ET_ADRESSE3').asstring;
   SetControlText('ETADRESSE123',St);
   SetControlText('ETCP',FormatCase(QSal.FindField('ET_CODEPOSTAL').asstring,'STR',5));
   SetControlText('ETVILLE',QSal.FindField('ET_VILLE').asstring);
   SetControlText('ETLIBELLE',QSal.FindField('ET_LIBELLE').asstring);
   SetControlText('ETTELEPHONE',FormatCase(QSal.FindField('ET_TELEPHONE').asstring,'STR',5));
   SetControlText('TELEPHONE',FormatCase(QSal.FindField('ET_TELEPHONE').asstring,'STR',5));//PT3-5
   SetControlText('ETTELECOPIE',FormatCase(QSal.FindField('ET_FAX').asstring,'STR',5));
   SetControlText('ETSIRET',FormatCase(QSal.FindField('ET_SIRET').asstring,'STR',5));
   SetControlText('ETAPE',FormatCase(QSal.FindField('ET_APE').asstring,'STR',5));

   if (QSal.FindField('PSA_CIVILITE').asstring='MR') or (QSal.FindField('PSA_CIVILITE').asstring='MM.') then
      Begin
      SetControlChecked('CKPSAM',True);     SetControlEnabled('CKPSAM',True);
      SetControlChecked('CKPSAMME',False);  SetControlEnabled('CKPSAMME',False);
      SetControlChecked('CKPSAMLLE',False); SetControlEnabled('CKPSAMLLE',False);
      End;
   if QSal.FindField('PSA_CIVILITE').asstring='MME' then
      Begin
      SetControlChecked('CKPSAM',False);    SetControlEnabled('CKPSAM',False);
      SetControlChecked('CKPSAMME',True);   SetControlEnabled('CKPSAMME',True);
      SetControlChecked('CKPSAMLLE',False); SetControlEnabled('CKPSAMLLE',False);
      End;
   if QSal.FindField('PSA_CIVILITE').asstring='MLE' then
      Begin
      SetControlChecked('CKPSAM',False);    SetControlEnabled('CKPSAM',False);
      SetControlChecked('CKPSAMME',False);  SetControlEnabled('CKPSAMME',False);
      SetControlChecked('CKPSAMLLE',True);  SetControlEnabled('CKPSAMLLE',True);
      End;
   if (QSal.FindField('PSA_CIVILITE').asstring='') or (QSal.FindField('PSA_CIVILITE').asstring='DR') then
      Begin
      SetControlChecked('CKPSAM',True);     SetControlEnabled('CKPSAM',True);
      SetControlChecked('CKPSAMME',True);   SetControlEnabled('CKPSAMME',True);
      SetControlChecked('CKPSAMLLE',True);  SetControlEnabled('CKPSAMLLE',True);
      End;

   //SetControlText('PSASALARIE',QSal.FindField('PSA_SALARIE').asstring);
   if Trim(QSal.FindField('PSA_NOMJF').asstring)='' then //DEB PT4-3
     Begin
     SetControlText('PSALIBELLE',QSal.FindField('PSA_LIBELLE').asstring);
     SetControlText('PSANOMFILLE',QSal.FindField('PSA_NOMJF').asstring);
     End
   else
     Begin
     SetControlText('PSALIBELLE',QSal.FindField('PSA_NOMJF').asstring);
     SetControlText('PSANOMFILLE',QSal.FindField('PSA_LIBELLE').asstring);
     End;                                         //FIN PT4-3
   SetControlText('PSAPRENOMS',QSal.FindField('PSA_PRENOM').asstring);
   SetControlText('PSASEXE',   QSal.FindField('PSA_SEXE').asstring);
   St:=QSal.FindField('PSA_ADRESSE1').asstring+' '+QSal.FindField('PSA_ADRESSE2').asstring+'  '+QSal.FindField('PSA_ADRESSE3').asstring;
   SetControlText('PSAADRESSE123',St);
   //SetControlText('PSAEMPLOI',RechDom('PGLIBEMPLOI',QSal.FindField('PSA_LIBELLEEMPLOI').asstring,False)); //PT5-3
   SetControlText('PSAEMPLOI',QSal.FindField('ET_ACTIVITE').AsString);
   SetControlText('PSACP',FormatCase(QSal.FindField('PSA_CODEPOSTAL').asstring,'STR',5));
   SetControlText('PSAVILLE',QSal.FindField('PSA_VILLE').asstring);
   SetControlText('PSASECU',FormatCase(QSal.FindField('PSA_NUMEROSS').asstring,'SS',5));
   if QSal.FindField('PSA_NATIONALITE').asstring<>'FRA' then
     SetControlText('PSANATIONALITE',RechDom('YYNATIONALITE',QSal.FindField('PSA_NATIONALITE').asstring,False)) //PT3-2 chgmt TTPAYS
   else
     SetControlText('PSANATIONALITE','');
   if QSal.FindField('PSA_NATIONALITE').asstring='FRA' then
     Begin
     SetControlChecked('CKFRANC',True);
     SetControlChecked('CKETRAN',False);
     End
   Else
     Begin
     SetControlChecked('CKFRANC',False);
     SetControlChecked('CKETRAN',True);
     End;
   SetControlText('PSADATENAIS',FormatCase(DateToStr(QSal.FindField('PSA_DATENAISSANCE').AsDateTime),'DATE',5));
   SetControlText('PSADPTNAIS',FormatCase(QSal.FindField('PSA_DEPTNAISSANCE').asstring,'STR',5));
   St:=QSal.FindField('PSA_COMMUNENAISS').asstring+'    '+RechDom('TTPAYS',QSal.FindField('PSA_PAYSNAISSANCE').asstring,False);
   SetControlText('PSALIEUNAIS',St);
   SetControlText('PSADATEEMBAUCHE',FormatCase(DateToStr(QSal.FindField('PSA_DATEENTREE').AsDateTime),'Date',5));
   Heure:='';
   DHeure:=QSal.FindField('PSA_HEUREMBAUCHE').AsDateTime;
   if  (DHeure=idate1900) then // or (DHeure=0)    PT19
     Heure:=''
   else DateTimeToString(Heure, 'hh:mm',QSal.FindField('PSA_HEUREMBAUCHE').AsDateTime);
     DecodeTime(DHeure, Hour, Min, Sec, MSec);
   SetControlText('PSAHREMBAUCHE',FormatCase(Heure,'HR',5));
//   SetControlText('PSAEMPLOI',RechDom('PGLIBEMPLOI',QSal.FindField('PSA_LIBELLEEMPLOI').asstring,False));//PT5-3

   SetControlText('LIEUFAIT',QSal.FindField('ET_VILLE').asstring);

   SetControlText('CODEMED','');     //DEB PT4-2
   SetControlText('LBCODEMED','');
   SetControlText('CPCORRESP','');   //FIN PT4-2
   { DEB PT14 Ajout Clause }
   //DEBUT PT18
   SetControlText('LIBELLEMED',QSal.FindField('ETB_MEDTRAVGU').asstring);
   st:=QSal.FindField('ANN_ALRUE1').asstring+' '+QSal.FindField('ANN_ALRUE2').asstring+' '+
       QSal.FindField('ANN_ALRUE3').asstring+' '+
       QSal.FindField('ANN_ALCP').asstring+' '+QSal.FindField('ANN_ALVILLE').asstring;
   SetControlText('ADRESSEMED',St);
   //FIN PT18
   END;
   Ferme(QSal);


end;

procedure TOF_PGDUE.InitTaille(ControlParent : TWinControl; taille : integer);
var
i : integer;
ChampTrouve:String;
ChampControl : TControl;
Edit : THEdit;
NumEdit:THNumEdit;
begin
     For i:=0 to ControlParent.ControlCount-1  do
        begin
           ChampControl:=ControlParent.Controls[i];
           ChampTrouve:=AnsiUpperCase(ChampControl.Name);
           if ChampControl is THEdit then
              begin
              Edit:=THEdit(ChampControl);
              Edit.Height:=taille;
              end;
           if ChampControl is THNumEdit then
              begin
              NumEdit:=THNumEdit(ChampControl);
              NumEdit.Height:=taille;
              end;
        end;
end;


procedure TOF_PGDUE.OnChange(Sender: TObject);
var
CheckOui,CheckNon : TCheckBox;
begin                         
//gestion des check Oui Non
{PT7 Mise en commentaire
CheckOui:=TCheckBox(GetControl('EXOOUI'));
CheckNon:=TCheckBox(GetControl('EXONON'));
GestionCheckOuiNon(CheckOui,CheckNon); }
CheckOui:=TCheckBox(GetControl('ABBATOUI'));
CheckNon:=TCheckBox(GetControl('ABBATNON'));
GestionCheckOuiNon(CheckOui,CheckNon);
CheckOui:=TCheckBox(GetControl('AGENTOUI'));
CheckNon:=TCheckBox(GetControl('AGENTNON'));
GestionCheckOuiNon(CheckOui,CheckNon);
CheckOui:=TCheckBox(GetControl('LICENOUI'));
CheckNon:=TCheckBox(GetControl('LICENNON'));
GestionCheckOuiNon(CheckOui,CheckNon);
CheckOui:=TCheckBox(GetControl('CBPREMIEROUI'));
CheckNon:=TCheckBox(GetControl('CBPREMIERNON'));
GestionCheckOuiNon(CheckOui,CheckNon);
SetControlEnabled('EFFECTIFAVEMB',(GetControlText('CBPREMIEROUI')='X'));
{PT7 Mise en commentaire
CheckOui:=TCheckBox(GetControl('SALFR'));
CheckNon:=TCheckBox(GetControl('SALEURO'));
GestionCheckOuiNon(CheckOui,CheckNon);}
end;


procedure TOF_PGDUE.DateElipsisclick(Sender: TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;


procedure TOF_PGDUE.InitDateFait(Sender: TObject);
begin
SetControlText('DATEFAIT',FormatCase(GetControlText('DATEFAITEDIT'),'DATE',5));
end;

procedure TOF_PGDUE.RensMedecine(Sender: TObject);
Var Q : TQuery;
    St : String;
begin
//DEBUT PT18
//DEB PT3-3
if GetControlText('LIBELLEMED')=''then
  Begin
   SetControlText('LBCODEMED','');  //PT4-2
   SetControlText('LIBELLEMED','');
   SetControlText('ADRESSEMED','');
   exit;
  End;
//FIN PT3-3
Q:=OpenSql('SELECT ANN_NOM1,ANN_ALRUE1,ANN_ALRUE2,ANN_ALRUE3,ANN_ALVILLE,ANN_ALCP '+
   'FROM ANNUAIRE WHERE ANN_GUIDPER="'+GetControlText('LIBELLEMED')+'" '+         //PT17
   'AND ANN_TYPEPER="MED" ',TRUE); //PT4-1
if not Q.eof Then
  Begin
   st:=Q.FindField('ANN_ALRUE1').asstring+' '+Q.FindField('ANN_ALRUE2').asstring+' '+
       Q.FindField('ANN_ALRUE3').asstring+' '+
       Q.FindField('ANN_ALCP').asstring+' '+Q.FindField('ANN_ALVILLE').asstring;
   SetControlText('ADRESSEMED',St);
  End
Else
  Begin
   SetControlText('ADRESSEMED','');
  End;
Ferme(Q);
//FIN PT18
end;
//DEB PT 2
procedure TOF_PGDUE.CPCorresp(Sender: TObject);
begin
SetControlText('CPCORRESP',FormatCase(GetControlText('CPCORRESP1'),'STR',5));
end;
//FIN PT 2
///PT3-1
procedure TOF_PGDUE.GestionCocheEmployeur(Sender: TObject);
var Ck : TcheckBox;
begin
if Sender is TCheckBox then Ck:=TCheckBox(Sender) else exit;
SetControlEnabled('CKETM',((Ck.Checked=False) AND (Ck.Name<>'CKETM')));
SetControlEnabled('CKETMME',((Ck.Checked=False) AND (Ck.Name<>'CKETMME')));
SetControlEnabled('CKETMLLE',((Ck.Checked=False) AND (Ck.Name<>'CKETMLLE')));
SetControlEnabled(Ck.name,True);
end;

//procedure TOF_PGDUE.ChangeCodeMed(Sender: TObject); DEB PT4-2
//begin
//SetControlText('LBCODEMED',FormatCase(GetControlText('CODEMED'),'STR',5));
//end;                                                FIN PT4-2
//PT3-4
procedure TOF_PGDUE.GestionTypeContrat(Sender: TObject);
var Ck : TcheckBox;
begin
if Sender is TCheckBox then Ck:=TCheckBox(Sender) else exit;
SetControlEnabled('CBDURIND',((Ck.Checked=False) AND (Ck.Name<>'CBDURIND')));
SetControlEnabled('CBDURDET',((Ck.Checked=False) AND (Ck.Name<>'CBDURDET')));
SetControlEnabled(Ck.name,True);
SetControlEnabled('DATEFINCONTRAT',GetControlText('CBDURDET')='X');
SetControlText('DATEFINCONTRAT','');
end;

procedure TOF_PGDUE.ControleValTablette(Sender:TObject);  //PT5-2
begin
If THEdit(Sender)=Nil Then Exit;
If (THEdit(Sender).Text<>'') and not (LookupValueExist(THEdit(Sender))) then
   begin
   If THEdit(Sender).Name='EMBAUCHE' Then
      begin  //PT6
      PGIBox('Le code saisi ne fait pas partie de la liste','Situation de l''employeur');
      SetFocusControl('EMBAUCHE');
      end;
   If THEdit(Sender).Name='QUALIF' Then
      begin //PT6
      PGIBox('Le code saisi ne fait pas partie de la liste','Nature de l''emploi');
      SetFocusControl('QUALIF');
      end;
   end;
end;

procedure TOF_PGDUE.ChangeEffAvEmb(Sender: TOBJect);
begin
SetControlText('EFFECTIFAV',FormatCase(GetControlText('EFFECTIFAVEMB'),'STR',5));
end;

procedure TOF_PGDUE.ChangeSalaireBrut(Sender: TOBJect);
begin
SetControlText('SALAIRE',FormatCase(GetControlText('SALAIREBRUT'),'NBR',5)); //PT12
end;

procedure TOF_PGDUE.ChangeEff(Sender: TOBJect);
begin
SetControlText('EFFECTIF',FormatCase(GetControlText('EFFECTIFSAIS'),'STR',5));  { PT13 }
end;
{ DEB PT16 }
procedure TOF_PGDUE.AffichDeclarant(Sender: TObject);
begin
if GetControlText('DECLARANT')='' then exit;
SetControlText('CONTACT'        ,RechDom('PGDECLARANTATTEST',GetControlText('DECLARANT'),False));
SetControlText('LIEUFAIT'       ,RechDom('PGDECLARANTVILLE' ,GetControlText('DECLARANT'),False));
SetControlText('EDITTELEPHONE'  ,RechDom('PGDECLARANTTEL' ,GetControlText('DECLARANT'),False));
SetControlText('TELEPHONE'      ,FormatCase(GetControltext('EDITTELEPHONE'),'STR',5));
end;
{ FIN PT16 }

{ DEB PT16 }
procedure TOF_PGDUE.AffichTelephone(Sender: TObject);
begin
SetControlText('TELEPHONE'      ,FormatCase(GetControltext('EDITTELEPHONE'),'STR',5));
end;
{ FIN PT16 }

Initialization
registerclasses([TOF_PGDUE]);
end.

