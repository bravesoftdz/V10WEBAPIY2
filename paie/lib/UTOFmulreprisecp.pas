{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 10/09/2001
Modifié le ... :   /  /
Description .. : Unit de saisie des reprises de CP
Mots clefs ... : PAIE;CP
*****************************************************************
PT1 26/11/2001 SB V563 Fiche de bug n°350
                       Modification de la requête des salaries à afficher
PT2 08/02/2002 PH V571 Correction idate1900
PT3 13/03/2002 SB V571 Fiche de bug n°10023 Date de sortie <= idate1900
PT4 04/02/2003 SB V591 FQ 10477 Affectation 0 code salarié
PT5 12/03/2004 SB V_50 FQ 11162 Encodage de de la date de cloture erroné si fin fevrier
PT6 07/04/2004 V_50 SB FQ 11212 Modification des critères d'affichage des salariés
PT7 09/04/2004 V_50 SB FQ 11136 Ajout Gestion des congés payés niveau salarié
PT8 08/10/2007 V_80 MF mise en place traitement des jours de fractionnement
}
unit UTOFMulRepriseCP;

interface
uses  StdCtrls,Controls,Classes,sysutils,
{$IFNDEF EAGLCLIENT}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,Fe_Main,Hqry,
{$ELSE}
      MaineAgl,HQry,emul,UTOB,
{$ENDIF}
      HCtrls,HEnt1,HMsgBox,UTOF,AglInit,ULibEditionPaie;

Type
     TOF_MulRepriseCP = Class (TOF)
       procedure OnArgument (stArgument: String);        Override;
       procedure GrilleDblClick (Sender: TObject);
       procedure OnLoad ; override;
       procedure ActiveWhere (Sender: TObject);
       procedure cknobullclick(Sender: TObject);
       procedure ExitEdit(Sender: TObject); //PT4
   private
       CP : boolean;
       etablissement : string;
       WW:THEdit ;
    END ;

implementation

uses PGoutils2,PgCalendrier,EntPaie;

procedure TOF_MulRepriseCP.OnArgument (stArgument: String);
var
{$IFNDEF EAGLCLIENT}
Grille:THDBGrid;
{$ELSE}
Grille:THGrid;
{$ENDIF}
//st : string;
cknobull  : TcheckBox;
Defaut    : ThEdit;
begin
Cp := false;
WW:=THEdit (GetControl ('XX_WHERE'));
cknobull      := TcheckBox(getcontrol('CKNOBULL'));
if cknobull   <> nil then
   begin
   cknobull.Checked := true;
   cknobull.OnClick := cknobullclick;
   end;

{$IFNDEF EAGLCLIENT}
Grille:=THDBGrid (GetControl ('Fliste'));
{$ELSE}
Grille:=THGrid (GetControl ('Fliste'));
{$ENDIF}
if Grille <> NIL then Grille.OnDblClick := GrilleDblClick;
{st := 'SELECT DISTINCT PSA_SALARIE FROM SALARIES WHERE '+
//      'PSA_SALARIE NOT IN (SELECT DISTINCT PPU_SALARIE FROM PAIEENCOURS) AND PSA_ETABLISSEMENT = "'+etablissement+'"'; //PT1 modif requête
      'PSA_SALARIE NOT IN (SELECT DISTINCT PCN_SALARIE FROM ABSENCESALARIE '+
                   'WHERE PCN_TYPEMVT="CPA" AND PCN_PERIODECP<2 '+
                'AND (((PCN_TYPECONGE="ACQ" OR PCN_TYPECONGE="ACA" '+
                       'OR PCN_TYPECONGE="ACS" OR PCN_TYPECONGE="REL" ) AND PCN_APAYES>0) '+
                      'OR PCN_TYPECONGE="ARR" OR (PCN_TYPECONGE="PRI" AND PCN_CODETAPE<>"...") '+
                      'OR PCN_TYPECONGE="AJP" OR PCN_TYPECONGE="AJU" OR PCN_TYPECONGE="SLD")) '+
      'AND PSA_ETABLISSEMENT = "'+etablissement+'"'; }
setControlVisible('CKNOBULL',False);
{ DEB PT4 }
Defaut:=ThEdit(getcontrol('PSA_SALARIE'));
If Defaut<>nil then Begin Defaut.MaxLength:=10; Defaut.OnExit:=ExitEdit; End;
{ FIN PT4 }
end;




procedure TOF_MulRepriseCP.OnLoad;
var
TEtablissement : THValCombobox;
begin
inherited;
ActiveWhere (NIL);
TEtablissement := THValCombobox(getcontrol('PSA_ETABLISSEMENT'));
if Tetablissement <> nil then
   if Tetablissement.value = '' then
           Tetablissement.Itemindex := 0;
end;

procedure TOF_MulRepriseCP.ActiveWhere(Sender: TObject);
var
TEtablissement                        : THValCombobox;
Q                                     : TQuery;
DateCloture,DateDebutN1,DateClotureN2 : tdatetime;
aa,mm,jj                              : word;
cknobull                              : TcheckBox;
SansBulletin                          : Boolean;
begin
inherited;
TEtablissement := THValCombobox(getcontrol('PSA_ETABLISSEMENT'));
if TEtablissement <> nil then
   begin
   if Tetablissement.value = '' then Tetablissement.Itemindex := 0;
   Q := Opensql('SELECT ETB_CONGESPAYES,ETB_DATECLOTURECPN FROM ETABCOMPL WHERE ETB_ETABLISSEMENT="'+Tetablissement.value+'"',True);
   if Not Q.EOF then //PORTAGECWAS
     Begin
     CP := (Q.findfield('ETB_CONGESPAYES').asstring = 'X');
     DateCloture := Q.findfield('ETB_DATECLOTURECPN').AsDateTime;
     end
    else
      DateCloture:=Idate1900;
   if not CP then // CP non gérés pour cet établissement
     begin
     PGIBox ('Traitement impossible : les congés payés ne sont pas gérés pour cet établissement', 'Reprise des congés payés');
     Tetablissement.Itemindex := 0;
     Ferme(Q);
     exit;
     end;
   Ferme(Q);
   if DateCloture<>idate1900 then
     Begin
     decodedate(DateCloture,aa,mm,jj);
     DateclotureN2:=PGEncodeDateBissextile(AA-2,MM,JJ);  { PT5 }
     DateDebutN1   := DateClotureN2+1;
     cknobull      := TcheckBox(getcontrol('CKNOBULL'));
     SansBulletin  := True;
     if cknobull <> nil then  SansBulletin := cknobull.Checked;
     if WW <> nil  then
       if SansBulletin then
//      WW.Text:=' PSA_SALARIE NOT IN (SELECT DISTINCT PPU_SALARIE FROM PAIEENCOURS) and psa_etablissement = "'+TEtablissement.value+'"'    PT1 Modif. ligne
        { PT6 Mise en commentaire remplacé par la clause suivante
        WW.Text:=' PSA_SALARIE NOT IN (SELECT DISTINCT PCN_SALARIE FROM ABSENCESALARIE '+
                'WHERE PCN_TYPEMVT="CPA" AND PCN_PERIODECP<2 '+
                'AND (((PCN_TYPECONGE="ACQ" OR PCN_TYPECONGE="ACA" '+
                       'OR PCN_TYPECONGE="ACS" OR PCN_TYPECONGE="REL" ) AND PCN_APAYES>0) '+
                      'OR PCN_TYPECONGE="ARR" OR (PCN_TYPECONGE="PRI" AND PCN_CODETAPE<>"...") '+
                      'OR PCN_TYPECONGE="AJP" OR PCN_TYPECONGE="AJU" OR PCN_TYPECONGE="SLD")) '+
// PT2 08/02/2002 PH V571 Correction idate1900
               'AND (PSA_DATESORTIE<="'+UsDateTime(iDate1900)+'" OR PSA_DATESORTIE IS NULL '+ //PT3 Ajout cond. =
               'OR PSA_DATESORTIE >= "'+usdatetime(DateDebutN1)+'")'                             }
        WW.Text:=' (PSA_DATESORTIE<="'+UsDateTime(iDate1900)+'"  OR PSA_DATESORTIE IS NULL '+
        'OR PSA_DATESORTIE >= "'+usdatetime(DateDebutN1)+'") AND PSA_CONGESPAYES="X" '+ { PT7 }
        'AND PSA_SALARIE NOT IN (SELECT DISTINCT ABS.PCN_SALARIE FROM ABSENCESALARIE ABS '+
        'WHERE ABS.PCN_TYPEMVT="CPA" '+
        'AND ( ( (ABS.PCN_TYPECONGE="ACQ" OR ABS.PCN_TYPECONGE="ACA" OR ABS.PCN_TYPECONGE="ACS" '+
                 'OR ABS.PCN_TYPECONGE="ACF" '+ // PT8
                 'OR ABS.PCN_TYPECONGE="REL" OR (ABS.PCN_TYPECONGE="AJP" AND ABS.PCN_SENSABS="+") '+
                                     'OR (ABS.PCN_TYPECONGE="AJU" AND ABS.PCN_SENSABS="+") ) '+
//PT8            'AND ABS.PCN_APAYES>0 AND ABS.PCN_PERIODECP=0) '+
                 'AND ABS.PCN_APAYES>0 AND (ABS.PCN_PERIODECP=0 OR ABS.PCN_PERIODECP=1)) '+
           'OR ( ((ABS.PCN_TYPECONGE="REP" AND ABS.PCN_APAYES>(SELECT SUM(ABS2.PCN_JOURS) '+
                                           'FROM ABSENCESALARIE ABS2 WHERE ABS2.PCN_SALARIE=ABS.PCN_SALARIE '+
                                           'AND ABS2.PCN_TYPECONGE="CPA" AND ABS2.PCN_TYPEMVT="CPA" '+
                                           'AND ABS2.PCN_PERIODECP=0)  ))  AND ABS.PCN_PERIODECP=0 ) '+
           'OR ( ((ABS.PCN_TYPECONGE="REP" AND ABS.PCN_APAYES>(SELECT SUM(ABS3.PCN_JOURS) '+
                                           'FROM ABSENCESALARIE ABS3 WHERE  ABS3.PCN_SALARIE=ABS.PCN_SALARIE '+
                                           'AND ABS3.PCN_TYPECONGE="CPA" AND ABS3.PCN_TYPEMVT="CPA" '+
                                           'AND ABS3.PCN_PERIODECP=1)  ))  AND ABS.PCN_PERIODECP=1 ) '+
           'OR (ABS.PCN_TYPECONGE="SLD" AND ABS.PCN_GENERECLOTURE="-" AND ABS.PCN_PERIODECP<2))) '
       else
         WW.Text:=' PSA_ETABLISSEMENT = "'+TEtablissement.value+'"'
         +' AND (PSA_DATESORTIE<="'+UsDateTime(iDate1900)+'" OR PSA_DATESORTIE IS NULL OR PSA_DATESORTIE >= "'+usdatetime(DateDebutN1)+'")';//PT3 Ajout cond. =
     end;
   end;

    // on recherche ts les salariés pour lesquels on n'a pas encore édité de bulletin
end;

procedure TOF_MulRepriseCP.GrilleDblClick(Sender: TObject);
var
cknobull       : TcheckBox;
Bull           : string;
begin
inherited;
ActiveWhere (NIL);
if not cp then
   begin
   PgiBox ('L''option de gestion des congés payés n''est pas cochée pour cet établissement','Reprise des mouvements CP impossible') ;
   exit;
   end;
bull := 'BULL' ;
cknobull := TcheckBox(getcontrol('CKNOBULL'));
if cknobull <> nil then
   begin
   if cknobull.checked then
      bull := 'SANS' ;
   end;

{$IFDEF EAGLCLIENT}
TheMulQ:=TOB (TFMul(Ecran).Q.TQ);
{$ELSE}
TheMulQ := THQuery(Ecran.FindComponent('Q'));
{$ENDIF}
AGLLanceFiche ('PAY','SAISREPRISECP',  '', '', GetControlText('PSA_ETABLISSEMENT')+';'+bull);
TheMulQ := NIL;
ActiveWhere (NIL);
end;


procedure TOF_MulRepriseCP.cknobullClick(Sender: TObject);
var
cknobull : TcheckBox;
init : word;
begin
cknobull := TcheckBox(getcontrol('CKNOBULL'));
if cknobull = nil then exit;
if cknobull.Checked = false then
   begin
   Init := HShowMessage('1;Reprise Congés payés;Attention, si vous décochez cette option, tous les salariés seront  affichés,#13#10y compris ceux pour lesquels un bulletin de paye a déjà été calculé !#13#10 Etes vous sûr de vouloir continuer ?;Q;YN;N;N;','','');
   if Init = mrYes then exit
   else
      cknobull.checked := true;
   end;
end;

{ DEB PT4 }
procedure TOF_MulRepriseCP.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;
{ FIN PT4 }

Initialization
registerclasses([TOF_MULRepriseCP]) ;

end.



