{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 15/06/2001
Modifié le ... :   /  /
Description .. : Multi critère de gestion déportée des absences
Suite ........ : Accède aux absences salariés et aux absences que le
Suite ........ : responsable doit valider
Mots clefs ... : PAIE;ABSENCES;PGDEPORTEE
*****************************************************************
PT1 SB 15/01/2002 Affichage de la liste des salariés Responsable et collezerodevant
PT2 SB 28/11/2002 V591 Intégration du calcul du recap en mode monobase ou base non déportée
PT3 SB 05/12/2002 V591 Ajout coche salariés sortis
PT4 SB 12/11/2003 V_42 Econges Ajout paramètres
PT5 SB 02/06/2005 V_60 FQ 12327 Econges : Gestion monobase => PCN_EXPORTOK
PT6 PH 20/06/2005 V_60 Mise en place de la hiérarchie
PT7 SB 12/07/2005 V_60 FQ 12231 Ajout message de confirmation
PT8 SB 25/07/2005 V_65 FQ 12399 Ajout récapitulatif pour affichage assistante
PT9 SB 25/07/2005 V_65 Rep PT6 Affichage des sous niveaux hierarchiques
PT10 PH 02/08/2005 V_65 FQ 12434
PT11 SB 05/12/2005 V_65 FQ 12737 Ajout sélection du niveau hierarchique
PT12 SB 16/11/2006 V_71 FQ 13690 Maj erroné des mvts exportés
}

unit UtofPgMulRecap;

interface
uses Controls,Classes,sysutils,HTB97,
{$IFDEF EAGLCLIENT}
      eMul,
{$ELSE}
      Mul,
{$ENDIF}
      HCtrls,HEnt1,HMsgBox,UTOF,
      stdctrls,paramsoc,PgOutils2,Ed_tools;

Type
     tof_PgMulRecap = Class (TOF)
       private
       WW    : THedit;
       Trait : String;
       procedure ActiveWhere (Sender: TObject);
       public
       procedure OnArgument(Arguments : String ) ; override ;
       procedure OnLoad; override;
       procedure OnExitSalarie(Sender : TObject);//PT1
       procedure ElipsisClickSal(Sender : TObject); //PT1
       procedure EcabCalculRecap(Sender : TObject); //PT2
       procedure ClickHierarchie(Sender: TObject);  { PT11 }
     END ;

implementation

uses Entpaie,P5Def,P5Util,pgoutilsEagl,PgHierarchie;

{ il convient de ne voir que les collaborateurs du responsable donc clause where basee sur la jointure
}
procedure tof_PgMulRecap.ActiveWhere(Sender: TObject);
var ChbxH : TCheckBox;
    Hierar : THMultiValComboBox;
    StHierar : String;
begin
ChbxH := TCheckBox (GetControl ('CKHIERARCHIE'));
Hierar := THMultiValComboBox (GetControl ('HIERARCHIE'));  { PT11 }
// initialisation du code salarié en fonction de la connection
WW.text:='';
{ DEB PT9 }
if Trait = 'RESP' then
  Begin
  If Assigned(ChbxH) and (ChbxH.Checked) and (LeSalarie <> '') then
      Begin
      if Assigned(Hierar) Then StHierar := RendClauseHierarchie(Hierar.Text); { PT11 }
      WW.text := StHierar + ' AND (PSE_RESPONSABS="'+LeSalarie+'" '+
                 'OR PSE_RESPONSABS IN (SELECT PGS_RESPONSABS FROM SERVICES,SERVICEORDRE '+
                 'WHERE PGS_CODESERVICE=PSO_CODESERVICE '+
                 'AND PSO_SERVICESUP IN (SELECT S.PGS_CODESERVICE FROM SERVICES S '+
                 'WHERE S.PGS_RESPONSABS="'+LeSalarie+'"))) ';
      End
  else
    if LeSalarie <> '' then  WW.text := ' AND PSE_RESPONSABS = "'+LeSalarie+'"'
    else WW.text := ' AND PSE_RESPONSABS = "0000000000" AND PSE_RESPONSABS <> "0000000000" ';
  End
else
  if Trait = 'ASS' then  { DEB PT8 }
     Begin
     If Assigned(ChbxH) and (ChbxH.Checked) and (LeSalarie <> '') then
     Begin
     if Assigned(Hierar) Then StHierar := RendClauseHierarchie(Hierar.Text);  { PT11 }
      WW.text := StHierar + ' AND (PSE_ASSISTABS="'+LeSalarie+'" '+
                'OR PSE_RESPONSABS IN (SELECT PGS_RESPONSABS FROM SERVICES,SERVICEORDRE '+
                'WHERE PGS_CODESERVICE=PSO_CODESERVICE '+
                'AND PSO_SERVICESUP IN (SELECT S.PGS_CODESERVICE FROM SERVICES S '+
                'WHERE S.PGS_SECRETAIREABS="'+LeSalarie+'"))) ';
      End
     else
     if LeSalarie <> '' then  WW.text := ' AND PSE_ASSISTABS = "'+LeSalarie+'"'
     else WW.text := ' AND PSE_ASSISTABS = "0000000000" AND PSE_ASSISTABS <> "0000000000" ';
     End;                { FIN PT8 }
{ FIN PT9 }
//DEB PT3
if GetControlText('CKSORTIE')='-' then
  Begin
  if WW.Text<>'' then WW.text:=WW.text+' AND';
   WW.text := WW.text+' AND (PSA_DATESORTIE<="'+UsDateTime(idate1900)+'" '+
                  'OR PSA_DATESORTIE is null '+
                  'OR PSA_DATESORTIE>="'+UsDatetime(Vh_Paie.PGECabDateIntegration)+'") ';
  End;

//FIN PT3
(*  PT-9 Mise en commentaire, Clause SQL érronée
ChbxH := TCheckBox (GetControl ('CKHIERARCHIE'));
  if ChbxH <> NIL then
    if ChbxH.Checked = True then
    Begin
    if Trait = 'RESP' then
      // DEB PT6
      begin
      WW.Text := WW.Text + 'OR (PSE_RESPONSABS IN (SELECT PGS_RESPONSVAR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'
              +' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSABS="'+LeSalarie+'")))';
      end;
      // FIN PT6
        *)

end;
{ DEB PT11 }
procedure tof_PgMulRecap.ClickHierarchie(Sender: TObject);
begin
SetcontrolEnabled('HIERARCHIE',(TCheckBox(Sender).Checked = True));
if TCheckBox(Sender).Checked = False then SetcontrolText('HIERARCHIE','');
end;
{ FIN PT11 }

procedure tof_PgMulRecap.EcabCalculRecap(Sender: TObject);
Var
St,Per,DebPer,FinPer : string;
begin
//DEB PT2 Calcul du récapitulatifs des salariés
if Trim(GetControlText('PRS_SALARIE'))<>'' then
  st:=' du salarié '+RechDom('PGSALARIE',Trim(GetControlText('PRS_SALARIE')),False)+'?'
else
  st:=' des salariés?';
If PgiAsk('Voulez-vous calculer le récapitulatif'+st,Ecran.Caption) =MrYes then
  Begin
  //Si monobase date d'intégration =Date de fin période en cours de l'exercice social
  RendPeriodeEnCours(Per,DebPer,FinPer);
  if (StrToDate(FinPer)>VH_Paie.PGECabDateIntegration) OR (VH_Paie.PGECabDateIntegration<=idate1900) then
    Begin
    SetParamSoc('SO_PGECABDATEINTEGRATION',StrToDate(FinPer));
    VH_Paie.PGECabDateIntegration:=StrToDate(FinPer);
    ChargeParamsPaie;
    End;
  { DEB PT5 }
  If PgiAsk('Le récapitulatif se calculera au '+DateToStr(VH_Paie.PGECabDateIntegration)+'. Voulez-vous continuer?',Ecran.caption) = MrNo then exit; { PT7 et PT10 }
  if VH_Paie.PGEcabMonoBase then
     Begin
     InitMoveProgressForm(nil,'Calcul des récapitulatifs','Traitement en cours, veuillez patienter...',2,False,True);
     MoveCurProgressForm;
     if GetControlText('PRS_SALARIE')<>'' then st := 'AND PCN_SALARIE="'+GetControlText('PRS_SALARIE')+'" ' else St := '';
     { DEB PT12 }
     ExecuteSql('UPDATE ABSENCESALARIE SET PCN_EXPORTOK = "X" '+
                'WHERE PCN_TYPEMVT="CPA" AND PCN_TYPECONGE="PRI" AND PCN_PERIODECP >= 2'+St);
     MoveCurProgressForm;
     { FIN PT12 }
     ExecuteSql('UPDATE ABSENCESALARIE SET PCN_EXPORTOK = "X" WHERE PCN_CODETAPE="P" '+
                'AND (PCN_TYPEMVT="ABS" OR (PCN_TYPECONGE="PRI" AND PCN_PERIODECP < 2)) '+  { PT12 }
                'AND PCN_TYPECONGE IN (SELECT PMA_MOTIFABSENCE FROM MOTIFABSENCE '+
                'WHERE ##PMA_PREDEFINI## PMA_MOTIFEAGL="X") '+St);
     MoveCurProgressForm;
     ExecuteSql('UPDATE ABSENCESALARIE SET PCN_EXPORTOK = "-" WHERE PCN_CODETAPE<>"P" '+
                'AND (PCN_TYPEMVT="ABS" OR (PCN_TYPECONGE="PRI" AND PCN_PERIODECP < 2)) '+  { PT12 }
                'AND PCN_TYPECONGE IN (SELECT PMA_MOTIFABSENCE FROM MOTIFABSENCE '+
                'WHERE ##PMA_PREDEFINI## PMA_MOTIFEAGL="X") '+St);
     FiniMoveProgressForm;
     End;
  { FIN PT5 }
  ChargeTobSalRecap('A',VH_Paie.PGECabDateIntegration,Trim(GetControlText('PRS_SALARIE')));
  CalculRecapAbsEnCours(Trim(GetControlText('PRS_SALARIE')));
  PGIInfo('Traitement terminé.',Ecran.Caption);
  TFMul(Ecran).BChercheClick(nil);
  End;
//FIN PT2
end;

procedure tof_PgMulRecap.ElipsisClickSal(Sender: TObject);
begin
If Trait<>'ADM' then AfficheTabSalResp(Sender,LeSalarie,Trait); { PT8 }

end;

procedure tof_PgMulRecap.OnArgument(Arguments: String);
var
Edit : THedit;
Btn : TToolBarButton97;
Num : Integer;
ChbxH : TCheckBox;
begin
inherited ;
WW:=THEdit (GetControl ('XX_WHERE'));
Trait := Trim (Arguments);
//DEB PT1
Edit := THedit(getcontrol('PRS_SALARIE'));
if Edit<>nil then
  begin Edit.OnExit :=OnExitSalarie;
  If Trait<>'ADM' then Edit.OnElipsisClick:=ElipsisClickSal;
  end;
//FIN PT1
//DEB PT2
SetControlVisible('BCALCULRECAP',(((VH_PAIE.PGEcabMonoBase) and (Trait='ADM')) OR (not VH_PAIE.PGEcabBaseDeporte)) and (Trait='ADM'));

{ DEB PT11 }
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
{ FIN PT11 }

Btn:=TToolBarButton97(GetControl('BCALCULRECAP'));
If Btn<>nil then Btn.OnClick:=EcabCalculRecap;
//FIN PT2
SetControlvisible('CKSORTIE',(Trait='ADM')); //PT3
{ DEB PT4 }
if ((Trait='ADM') OR (Trait='RESP')) AND (GetControl ('PSA_TRAVAILN1')<>nil) then
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
{ FIN PT4 }

end;

//DEB PT1
procedure tof_PgMulRecap.OnExitSalarie(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;

procedure tof_PgMulRecap.OnLoad;
begin
inherited;
ActiveWhere (NIL);
if Trait = 'RESP' then TFmul(Ecran).caption := 'Récapitulatif absences des salariés du responsable '+LeNomSal
 else if Trait = 'ASS' then TFmul(Ecran).caption := 'Récapitulatif absences des salariés liés à l''assistant(e) '+LeNomSal { PT8 } 
 else TFmul(Ecran).caption := 'Administrateur : Récapitulatifs des absences';
UpdateCaption(TFmul(Ecran));
end;
//FIN PT1

Initialization
registerclasses([tof_PgMulRecap]);
end.
