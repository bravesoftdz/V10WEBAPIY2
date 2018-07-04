{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 25/04/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGEDITCOMPTEURSDIF ()
Mots clefs ... : TOF;PGEDITCOMPTEURSDIF
*****************************************************************}
{
PT1 26/01/2007  V_80 FC Mise en place du filtage habilitation pour les lookuplist
                        pour les critères code salarié uniquement
PF2 08/02/2007  V_80 FC Suppression de l'espace entre YEAR et ( car non compatible Oracle
}
Unit UTofPGEditCompteursDIF ;

Interface

uses StdCtrls,
     Controls,
     Classes,
     forms,
     sysutils,
     ComCtrls,
{$IFDEF EAGLCLIENT}
     eQRS1,
     UtileAgl,
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     EdtEtat,
     QRS1,
     EdtREtat,
{$ENDIF}
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     ParamSoc,
     UTOB,
     HTB97,
     PgEditOutils,
     PgEditOutils2,
     EntPaie,
     PGoutils,
     PGoutils2,
     PGEdtEtat,
     LookUp,
     HStatus,
     UTobDebug,
     windows,
     HQry,
     ed_tools;

Type
  TOF_PGEDITCOMPTEURSDIF = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure OnUpdate;              override;
    private
    TobEditDIF : Tob;
    {$IFDEF EMANAGER}
    TypeUtilisat : String;
    MultiNiveau : Boolean;
    {$ENDIF}
//    procedure LancerDIF(Sender : TObject);
    procedure RemplirDroits(StWhere : String);
    procedure RemplirFormations(StWhere : String);
    Procedure RemplirDemande(StWhere : String);
    procedure ExitEdit(Sender: TObject);
    {$IFDEF EMANAGER}
    procedure SalarieElipsisClick(Sender : Tobject);
    procedure RespElipsisClick(Sender : TObject);
    {$ENDIF}
  end ;

Implementation

procedure TOF_PGEDITCOMPTEURSDIF.OnUpdate;
var Pages : TPageControl;
    Where : String;
begin
Inherited ;
  Pages := TPageControl(GetControl('PAGES'));
  Where := RecupWhereCritere(Pages);
  If GetCheckBoxState('CKSORTIE') = Cbchecked then
  begin
       If Where <> '' then Where := Where + ' AND (PSA_DATESORTIE IS NULL OR PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'" OR PSA_DATESORTIE>="'+UsDateTime(StrToDate(GetControlText('DATEARRET')))+'")'
       else Where := 'WHERE (PSA_DATESORTIE IS NULL OR PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'" OR PSA_DATESORTIE>="'+UsDateTime(StrToDate(GetControlText('DATEARRET')))+'")';
  end;
  TobEditDIF := Tob.Create('EditionDIF',Nil,-1);
  InitMoveProgressForm(nil, 'Chargement des données', 'Veuillez patienter SVP ...', 0, FALSE, TRUE);
  InitMove(0, '');
  {$IFDEF EMANAGER}
  MultiNiveau := GetCheckBoxState('CMULTINIVEAU')= CbChecked;
  {$ENDIF}
  RemplirDroits(Where);
  RemplirFormations(Where);
  RemplirDemande(Where);
  FiniMoveProgressForm;
  TobEditDIF.Detail.Sort('PSA_SALARIE;ANNEE;DATEDEB');
  TFQRS1(Ecran).LaTob:= TobEditDIF;
end ;

procedure TOF_PGEDITCOMPTEURSDIF.OnClose;
begin
  Inherited ;
  FreeAndNil (TobEditDIF);
end;

{procedure TOF_PGEDITCOMPTEURSDIF.LancerDIF(Sender : TObject);
var Pages : TPageControl;
    Where : String;
begin
  Pages := TPageControl(GetControl('PAGES'));
  Where := RecupWhereCritere(Pages);
  If GetCheckBoxState('CKSORTIE') = Cbchecked then
  begin
       If Where <> '' then Where := Where + ' AND (PSA_DATESORTIE IS NULL OR PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'" OR PSA_DATESORTIE>="'+UsDateTime(StrToDate(GetControlText('DATEARRET')))+'")'
       else Where := 'WHERE (PSA_DATESORTIE IS NULL OR PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'" OR PSA_DATESORTIE>="'+UsDateTime(StrToDate(GetControlText('DATEARRET')))+'")';
  end;
  TobEditDIF := Tob.Create('EditionDIF',Nil,-1);
  InitMoveProgressForm(nil, 'Chargement des données', 'Veuillez patienter SVP ...', 0, FALSE, TRUE);
  InitMove(0, '');}
  {$IFDEF EMANAGER}
//  MultiNiveau := GetCheckBoxState('CMULTINIVEAU')= CbChecked;
  {$ENDIF}
{  RemplirDroits(Where);
  RemplirFormations(Where);
  RemplirDemande(Where);
  FiniMoveProgressForm;
  TobEditDIF.Detail.Sort('PSA_SALARIE;ANNEE;DATEDEB');
end ;}

procedure TOF_PGEDITCOMPTEURSDIF.OnArgument (S : String ) ;
var //BVal : TToolBarButton97;
    Min,Max : String;
    Defaut: THEdit;
begin
  Inherited ;
  SetControlText('DATEARRET',DateToStr(V_PGI.DateEntree));
//  BVal := TToolBarButton97(GetControl('BLANCEDIF'));
//  If BVal <> Nil then BVal.OnClick := LancerDIF;
  SetControlVisible('BVALIDER',True);
  SetControlVisible('BLANCEDIF',False);
  SetControlText('DOSSIER',GetParamSoc ('SO_LIBELLE'));

  RecupMinMaxTablette('PG', 'ETABLISS', 'ET_ETABLISSEMENT', Min, Max);
  SetControlText('PSA_ETABLISSEMENT', Min);
  SetControlText('PSA_ETABLISSEMENT_', Max);
  SetControltext('PSA_DATEENTREE',DateToStr(IDate1900));
  SetControltext('PSA_DATEENTREE_',DateToStr(V_PGI.DateEntree));
  SetControlChecked('CKSORTIE',True);
  Defaut:=ThEdit(getcontrol('PSA_SALARIE'));
  If Defaut<>nil then Defaut.OnExit:=ExitEdit;
  Defaut:=ThEdit(getcontrol('PSA_SALARIE_'));
  If Defaut<>nil then Defaut.OnExit:=ExitEdit;
  {$IFDEF EMANAGER}
  SetControlVisible('CMULTINIVEAU',True);
  Defaut:=ThEdit(getcontrol('PSA_SALARIE'));
  If Defaut<>nil then Defaut.OnElipsisClick := SalarieElipsisClick;
  Defaut:=ThEdit(getcontrol('PSA_SALARIE_'));
  If Defaut<>nil then Defaut.OnElipsisClick:= SalarieElipsisClick;
  SetControlProperty('Option','TabVisible',False);
  If ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"') then TypeUtilisat := 'R'
  else TypeUtilisat := 'S';
  If Defaut<>nil then Defaut.OnElipsisClick := SalarieElipsisClick;
  SetControlVisible('PSE_RESPONSFOR',True);
  SetControlVisible('TPSE_RESPONSFOR',True);
  SetControlVisible('LIBRESP',True);
  SetControlCaption('LIBRESP','');
  Defaut:=ThEdit(getcontrol('PSE_RESPONSFOR'));
  If Defaut<>nil then Defaut.OnElipsisClick := RespElipsisClick;
  {$ELSE}
  RecupMinMaxTablette('PG','SALARIES','PSA_SALARIE',Min,Max);
  SetControlText('PSA_SALARIE',Min);
  SetControlText('PSA_SALARIE_',Max);
  {$ENDIF}
end ;

procedure TOF_PGEDITCOMPTEURSDIF.RemplirDroits(StWhere : String);
var T,TCumul : tob;
    Cumul : String;
    Q : TQuery;
    i,AnneeI : Integer;
    aa, mm, dd : Word;
    DateArret : TDateTime;
    Salarie : String;
    LeWhere : String;
begin
     DateArret := StrToDate(GetControlText('DATEARRET'));
     If StWhere <> '' then LeWhere := StWhere + ' AND PHC_CUMULPAIE="'+Cumul+'"'
     else LeWhere := 'WHERE PHC_CUMULPAIE="'+Cumul+'"';
     Cumul := GetParamSOc('SO_PGCUMULDIFACQUIS');
     {$IFDEF EMANAGER}
     If typeUtilisat = 'R' then
     begin
          If MultiNiveau then LeWhere := LeWhere + ' AND (PSE_RESPONSFOR="'+V_PGI.UserSalarie+'" OR PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
          ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))'
          else LeWhere := LeWhere + ' AND PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"';
     end
      else LeWhere := LeWhere + ' AND PSE_RESPONSFOR IN '+
     '(SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
     'WHERE (PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")))';
     {$ENDIF}
     LeWhere := LeWhere + ' GROUP BY PHC_SALARIE,YEAR(PHC_DATEDEBUT)'; //PT2
     Q := OpenSQL('SELECT PHC_SALARIE,YEAR(PHC_DATEDEBUT) ANNEE,SUM(PHC_MONTANT) ACQUIS '+
     'FROM HISTOCUMSAL '+
     'LEFT JOIN SALARIES ON PSA_SALARIE=PHC_SALARIE '+
     'LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE '+LeWhere,True);
     TCumul := Tob.Create('LesCumuls',Nil,-1);
     TCumul.LoadDetailDB('LesCumuls','','',Q,False);
     Ferme(Q);
     For i := 0 To TCumul.Detail.Count -1 do
     begin
          Salarie := TCumul.Detail[i].GetValue('PHC_SALARIE'); //@@@@
          AnneeI := StrToInt(TCumul.Detail[i].GetValue('ANNEE'));
          T := Tob.Create('FilleEdition',TobEditDIF,-1);
          T.AddChampSupValeur('PSA_SALARIE',Salarie);
          T.AddChampSupValeur('TYPE','DROITS');
          decodedate(DateArret, aa, mm, dd);
          If aa <> AnneeI then
          begin
               T.AddChampSupValeur('LIBELLE','Cumul au '  + DateToStr(EncodeDate(AnneeI,12,31)));
               T.AddChampSupValeur('ANNEE',TCumul.Detail[i].GetValue('ANNEE'));
               T.AddChampSupValeur('DATEDEB',EncodeDate(AnneeI,12,31));
               T.AddChampSupValeur('DATEFIN',EncodeDate(AnneeI,12,31));
               T.AddChampSupValeur('DRACQ',TCumul.Detail[i].GetValue('ACQUIS'));
               T.AddChampSupValeur('DREC',0);

          end
          else
          begin
               T.AddChampSupValeur('LIBELLE','Cumul au '  + DateToStr(FindeMois(PlusMois(DateArret,-1))));
               T.AddChampSupValeur('ANNEE',TCumul.Detail[i].GetValue('ANNEE'));
               T.AddChampSupValeur('DATEDEB',FindeMois(PlusMois(DateArret,-1)));
               T.AddChampSupValeur('DATEFIN',FinDeMois(PlusMois(DateArret,-1)));
               T.AddChampSupValeur('DRACQ',0);
               T.AddChampSupValeur('DREC',TCumul.Detail[i].GetValue('ACQUIS'));
          end;
          T.AddChampSupValeur('FTT',0);
          T.AddChampSupValeur('FHTT',0);
          T.AddChampSupValeur('FTOTAL',0);
          T.AddChampSupValeur('DETT',0);
          T.AddChampSupValeur('DEHTT',0);
          T.AddChampSupValeur('DETOTAL',0);
          T.AddChampSupValeur('PFI_ETATINSCFOR','');
          T.AddChampSupValeur('SOLDE',0);
          T.AddChampSupValeur('SOLDETHEO',0);
     end;
     TCumul.Free;

end;

procedure TOF_PGEDITCOMPTEURSDIF.RemplirFormations(StWhere : String);
var TobLesformations,T : Tob;
    i : Integer;
    Q : TQuery;
    DateArret,DF,DD : TDateTime;
    Salarie,Libelle : String;
    LeWhere : String;
begin
     DateArret := StrToDate(GetControlText('DATEARRET'));
     If StWhere <> '' then LeWhere := StWhere + ' AND PFO_DATEFIN<="'+UsDateTime(DateArret)+'" AND PFO_TYPEPLANPREV="DIF"'
     else LeWhere := 'WHERE PFO_DATEFIN<="'+UsDateTime(DateArret)+'" AND PFO_TYPEPLANPREV="DIF"';
     {$IFDEF EMANAGER}
     If typeUtilisat = 'R' then
     begin
          If MultiNiveau then LeWhere := LeWhere + ' AND (PSE_RESPONSFOR="'+V_PGI.UserSalarie+'" OR PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
          ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))'
          else LeWhere := LeWhere + ' AND PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"';
     end
      else LeWhere := LeWhere + ' AND PSE_RESPONSFOR IN '+
     '(SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
     'WHERE (PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")))';
     {$ENDIF}
     Q := OpenSQL('SELECT PFO_SALARIE,YEAR(PFO_DATEFIN) ANNEE,PFO_DATEDEBUT,PFO_DATEFIN,PFO_CODESTAGE,'+
     'PFO_HTPSNONTRAV,PFO_HTPSTRAV,PFO_NBREHEURE FROM FORMATIONS '+
     'LEFT JOIN SALARIES ON PSA_SALARIE=PFO_SALARIE '+
     'LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE '+LeWhere,True);
     TobLesformations := Tob.Create('Lesformations',Nil,-1);
     TobLesformations.LoadDetailDB('Lesformations','','',Q,False);
     Ferme(Q);
     For i := 0 To TobLesformations.Detail.Count -1 do
     begin
          Salarie := TobLesformations.Detail[i].GetValue('PFO_SALARIE');
          T := Tob.Create('FilleEdition',TobEditDIF,-1);
          T.AddChampSupValeur('PSA_SALARIE',Salarie);
          T.AddChampSupValeur('TYPE','FOR');
          DD := TobLesformations.Detail[i].GetValue('PFO_DATEDEBUT');
          DF := TobLesformations.Detail[i].GetValue('PFO_DATEFIN');
          Libelle := 'Formation '+ RechDom('PGSTAGEFORMCOMPLET',TobLesformations.Detail[i].GetValue('PFO_CODESTAGE'),False)+ ' du ' +
          formatdateTime('dd/mm/yy',DD) + ' au '+FormatDateTime('dd/mm/yy',DF);
          T.AddChampSupValeur('LIBELLE',Libelle);
          T.AddChampSupValeur('ANNEE',TobLesformations.Detail[i].GetValue('ANNEE'));
          T.AddChampSupValeur('DATEDEB',TobLesformations.Detail[i].GetValue('PFO_DATEDEBUT'));
          T.AddChampSupValeur('DATEFIN',TobLesformations.Detail[i].GetValue('PFO_DATEFIN'));
          T.AddChampSupValeur('DRACQ',0);
          T.AddChampSupValeur('DREC',0);
          T.AddChampSupValeur('FTT',TobLesformations.Detail[i].GetValue('PFO_HTPSTRAV'));
          T.AddChampSupValeur('FHTT',TobLesformations.Detail[i].GetValue('PFO_HTPSNONTRAV'));
          T.AddChampSupValeur('FTOTAL',TobLesformations.Detail[i].GetValue('PFO_NBREHEURE'));
          T.AddChampSupValeur('DETT',0);
          T.AddChampSupValeur('DEHTT',0);
          T.AddChampSupValeur('DETOTAL',0);
          T.AddChampSupValeur('PFI_ETATINSCFOR','');
          T.AddChampSupValeur('SOLDE',0);
          T.AddChampSupValeur('SOLDETHEO',0);
     end;
     TobLesformations.Free;
end;

Procedure TOF_PGEDITCOMPTEURSDIF.RemplirDemande(StWhere : String);
var TobLesformations,T : Tob;
    i : Integer;
    Q : TQuery;
    Salarie : String;
    LeWhere : String;
begin
     If StWhere <> '' then LeWhere := StWhere + ' AND (PFI_ETATINSCFOR="ATT" OR PFI_ETATINSCFOR="VAL") AND PFI_TYPEPLANPREV="DIF" AND PFI_REALISE="-"'
     else LeWhere := 'WHERE (PFI_ETATINSCFOR="ATT" OR PFI_ETATINSCFOR="VAL") AND PFI_TYPEPLANPREV="DIF" AND PFI_REALISE="-"';
     {$IFDEF EMANAGER}
     If typeUtilisat = 'R' then
     begin
          If MultiNiveau then LeWhere := LeWhere + ' AND (PSE_RESPONSFOR="'+V_PGI.UserSalarie+'" OR PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
          ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))'
          else LeWhere := LeWhere + ' AND PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"';
     end
      else LeWhere := LeWhere + ' AND PSE_RESPONSFOR IN '+
     '(SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
     'WHERE (PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")))';
     {$ENDIF}
     Q := OpenSQL('SELECT PFI_DATEDIF,PFI_SALARIE,PFI_CODESTAGE,PFI_MILLESIME,PFI_ETATINSCFOR,'+
     'PFI_HTPSNONTRAV,PFI_HTPSTRAV,PFI_DUREESTAGE FROM INSCFORMATION '+
     'LEFT JOIN SALARIES ON PSA_SALARIE=PFI_SALARIE '+
     'LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE '+LeWhere,True);
     TobLesformations := Tob.Create('Lesformations',Nil,-1);
     TobLesformations.LoadDetailDB('Lesformations','','',Q,False);
     Ferme(Q);
     For i := 0 To TobLesformations.Detail.Count -1 do
     begin
          Salarie := TobLesformations.Detail[i].GetValue('PFI_SALARIE');
          T := Tob.Create('FilleEdition',TobEditDIF,-1);
          T.AddChampSupValeur('PSA_SALARIE',Salarie);
          T.AddChampSupValeur('TYPE','DEMANDE');
          T.AddChampSupValeur('LIBELLE',RechDom('PGSTAGEFORMCOMPLET',TobLesformations.Detail[i].GetValue('PFI_CODESTAGE'),False));
          T.AddChampSupValeur('ANNEE',TobLesformations.Detail[i].GetValue('PFI_MILLESIME'));
          T.AddChampSupValeur('DATEDEB',TobLesformations.Detail[i].GetValue('PFI_DATEDIF'));
          T.AddChampSupValeur('DATEFIN',TobLesformations.Detail[i].GetValue('PFI_DATEDIF'));
          T.AddChampSupValeur('DRACQ',0);
          T.AddChampSupValeur('DREC',0);
          T.AddChampSupValeur('FTT',0);
          T.AddChampSupValeur('FHTT',0);
          T.AddChampSupValeur('FTOTAL',0);
          T.AddChampSupValeur('DETT',TobLesformations.Detail[i].GetValue('PFI_HTPSTRAV'));
          T.AddChampSupValeur('DEHTT',TobLesformations.Detail[i].GetValue('PFI_HTPSNONTRAV'));
          T.AddChampSupValeur('DETOTAL',TobLesformations.Detail[i].GetValue('PFI_DUREESTAGE'));
          T.AddChampSupValeur('PFI_ETATINSCFOR',TobLesformations.Detail[i].GetValue('PFI_ETATINSCFOR'));
          T.AddChampSupValeur('SOLDE',0);
          T.AddChampSupValeur('SOLDETHEO',0);
     end;
     TobLesformations.Free;
end;

procedure TOF_PGEDITCOMPTEURSDIF.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;

{$IFDEF EMANAGER}
procedure TOF_PGEDITCOMPTEURSDIF.SalarieElipsisClick(Sender : Tobject);
var StWhere,StOrder : String;
begin
     If typeUtilisat = 'R' then StWhere := 'PSE_CODESERVICE IN '+
     '(SELECT PGS_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
     ' WHERE (PSE_RESPONSFOR="'+V_PGI.UserSalarie+'" AND (PSO_NIVEAUSUP=0 OR PSO_NIVEAUSUP=1))'+
     ' OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'"))'
     else StWhere := 'PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
     'WHERE (PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")))';
     If GetControlText('RESPONSFOR') <> '' then StWhere := StWhere + ' AND PSE_RESPONSFOR="'+GetControlText('RESPONSFOR')+'"';
     If (TypeUtilisat = 'R') and (GetCheckBoxState('CMULTINIVEAU') <> CbChecked) then StWhere := 'PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"';
     StOrder := 'PSA_SALARIE';
     StWhere := RecupClauseHabilitationLookupList(StWhere);  //PT1
     LookupList(THEdit(Sender),'Liste des salariés','SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE','PSA_SALARIE','PSA_LIBELLE,PSA_PRENOM',StWhere,StOrder, True,-1);
end;

procedure TOF_PGEDITCOMPTEURSDIF.RespElipsisClick(Sender : TObject);
var StOrder,StWhere : String;
begin
         If TypeUtilisat = 'R' then StWhere := '(PSI_TYPEINTERIM="SAL" OR PSI_TYPEINTERIM="EXT") AND (PSI_INTERIMAIRE="'+V_PGI.UserSalarie+'" OR PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
            ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))'
        else StWhere := '(PSI_TYPEINTERIM="SAL" OR PSI_TYPEINTERIM="EXT") AND PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '+
        'WHERE (PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")))';
        StOrder := 'PSI_LIBELLE';
        LookupList(THEdit(Sender),'Liste des responsables','INTERIMAIRES','PSI_INTERIMAIRE','PSI_LIBELLE,PSI_PRENOM',StWhere,StOrder, True,-1);
end;

{$ENDIF}


Initialization
  registerclasses ( [ TOF_PGEDITCOMPTEURSDIF ] ) ;
end.



