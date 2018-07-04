{***********UNITE*************************************************
Auteur  ...... : PAIE PGI GGU
Créé le ...... : 04/01/2008
Modifié le ... :   /  /
Description .. : Edition des compétences à acquérir pour le poste exercé et des formations liées
Mots clefs ... : PAIE;COMPETENCES;FORMATION
*****************************************************************
 }
unit UTOF_PGEDITCPTCESACQUER;

interface
uses StdCtrls, Controls, Classes,sysutils, ComCtrls,
  {$IFDEF EAGLCLIENT}
  eQRS1, 
  {$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} QRS1,
  {$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOF, ParamDat,
  ParamSoc, HQry, utob;

type

  TOF_PGEDITCPTCESACQUER = class(TOF)
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnUpdate; override;
    procedure OnClose; override;
  private
    TobAEditer : Tob;
  end;
  
implementation

uses
  PGTobOutils, ed_Tools, StrUtils;

procedure TOF_PGEDITCPTCESACQUER.OnArgument(Arguments: string);
begin
  inherited;
  SetControlText('DATE', AGLDateToStr(Date()));
end;

procedure TOF_PGEDITCPTCESACQUER.OnUpdate;
var
  TempTob, TobCompetForm, TempTobFormation : Tob;
  TempTobSal, TempTobLigneEdition : Tob;
  TobSalarieCompetManq, TempTobSalCompet : Tob;
  WhereUtilisateur, params, param : String;
  i, j : Integer;
  Procedure AddTobChampsLibresRH(var LaTob : Tob);
  var
    i : Integer;
  begin
    { On récupère les noms des champs libres }
    For i := 1 to 3 do
    begin
      LaTob.AddChampSupValeur('TABLELIBRECR'+IntToStr(i),RechDom('GCZONELIBRE','CP'+IntToStr(i),False));
    end;
    For i := 1 to 5 do
    begin
      LaTob.AddChampSupValeur('LIBRERH'+IntToStr(i),RechDom('GCZONELIBRE','RH'+IntToStr(i),False));
    end;
    TobAEditer.AddChampSupValeur('SO_PGFFORLIBRE1', GetParamSoc('SO_PGFFORLIBRE1'));
    TobAEditer.AddChampSupValeur('SO_PGFFORLIBRE2', GetParamSoc('SO_PGFFORLIBRE2'));
    TobAEditer.AddChampSupValeur('SO_PGFFORLIBRE3', GetParamSoc('SO_PGFFORLIBRE3'));
    TobAEditer.AddChampSupValeur('SO_PGFFORLIBRE4', GetParamSoc('SO_PGFFORLIBRE4'));
    TobAEditer.AddChampSupValeur('SO_PGFFORLIBRE5', GetParamSoc('SO_PGFFORLIBRE5'));
    TobAEditer.AddChampSupValeur('SO_PGFFORLIBRE6', GetParamSoc('SO_PGFFORLIBRE6'));
    TobAEditer.AddChampSupValeur('SO_PGFFORLIBRE7', GetParamSoc('SO_PGFFORLIBRE7'));
    TobAEditer.AddChampSupValeur('SO_PGFFORLIBRE8', GetParamSoc('SO_PGFFORLIBRE8'));
    TobAEditer.AddChampSupValeur('CRI_EXPORT', GetControlText('FListe'));
  end;
begin
  inherited;
  { On récupère les saisies utilisateur }
  if (GetControlText('PSA_ETABLISSEMENT') <> '') AND (Pos('>', GetControlText('PSA_ETABLISSEMENT')) = 0) then
  begin
    WhereUtilisateur := ' AND PSA_ETABLISSEMENT IN (';
    params := GetControlText('PSA_ETABLISSEMENT');
    param := READTOKENST(params);
    While Param <> '' do
    begin
      WhereUtilisateur := WhereUtilisateur + '"'+Param+'",';
      param := READTOKENST(params);
    end;
    WhereUtilisateur := LeftStr(WhereUtilisateur, Length(WhereUtilisateur) -1);
    WhereUtilisateur := WhereUtilisateur +') ';
  end;
  if (GetControlText('PSA_SALARIE') <> '') AND (GetControlText('PSA_SALARIE_') <> '') then
    WhereUtilisateur := WhereUtilisateur + ' AND PSA_SALARIE BETWEEN '+GetControlText('PSA_SALARIE')+' AND '+GetControlText('PSA_SALARIE_')+' ';
  If Assigned(TobAEditer) then FreeAndNil(TobAEditer);
  TobAEditer := Tob.Create('Données de l''édition', nil, -1);

//Liste des compétences à acquérir pour tenir l'emploi actuel :
  TempTob := Tob.Create('Liste des salariés et des compétences manquantes', nil, -1);
  TempTob.LoadDetailFromSQL(
    'select PSA_SALARIE, PSA_LIBELLE, PSA_PRENOM, PSA_LIBELLEEMPLOI, PSA_ETABLISSEMENT, PSA_DATEENTREE, '
    +'PSA_TRAVAILN1, PSA_TRAVAILN2, PSA_TRAVAILN3, PSA_TRAVAILN4, CC_LIBELLE as LIBELLEEMPLOI, '
    +'POS_COMPETENCE, POS_LIBELLE, POS_TABLELIBRECR1, POS_TABLELIBRECR2, POS_TABLELIBRECR3, POS_DEGREMAITRISE, '
    +'PCO_TABLELIBRERH1, PCO_TABLELIBRERH2, PCO_TABLELIBRERH3, PCO_TABLELIBRERH4, PCO_TABLELIBRERH5 '
    //Liste des salariés
    +'from salaries s1 '
    //recherche du libellé de l'emploi
    +'left join CHOIXCOD '
    +'on ((CC_TYPE="PLE") and (CC_CODE = s1.PSA_LIBELLEEMPLOI)) '
    //liste des compétences nécéssaires à l'emploi
    +'left join STAGEOBJECTIF '
    +'on ((pos_codeStage = psa_libelleemploi) and (pos_natobjstage = "EMP")) '
    //Propriétés des compétences
    +'LEFT JOIN RHCOMPETENCES '
    +'ON (POS_COMPETENCE=PCO_COMPETENCE) '
    //On enlève les compétences que le salarié possède déjà pour n'avoir que les compétences à acquérir
    +'WHERE pos_competence not in ( '
    +'    select pch_competence '
    +'    from salaries s2 '
    +'    left join PGRHCOMPETRESSOURCE '
    +'    on ((pch_salarie = s2.psa_salarie) and (PCH_TYPERESSOURCE = "SAL")) '
    +'    where s1.psa_salarie = pch_salarie '
    //On ne prends en compte que les compétences dont le degré de maitrise est supérieur ou égal au degré nécéssaire au poste
    +'      and pos_degremaitrise <= pch_degremaitrise '
    //On ne prends en compte que les compétences valides au jour choisit
    +'      and pch_datevalidation <= "'+ USDATETIME(AglStrToDate(GetControlText('DATE'))) +'" '
    +'      and ((pch_datefinvalid = "19000101 00:00:00") or (pch_datefinvalid >= "'+ USDATETIME(AglStrToDate(GetControlText('DATE'))) +'")) '
    +') '
    +'AND POS_COMPETENCE is not null '+ WhereUtilisateur);
  if Assigned(TobSalarieCompetManq) then FreeAndNil(TobSalarieCompetManq);
  EclateTob(TempTob, TobSalarieCompetManq, 'PSA_SALARIE', False);
  FreeAndNil(TempTob);
  { La tob mère doit contenir tous les champs utilisés par l'état }
  { Salarié }
  TobAEditer.AddChampSup('PSA_SALARIE'       ,False);
  TobAEditer.AddChampSup('PSA_LIBELLE'       ,False);
  TobAEditer.AddChampSup('PSA_PRENOM'        ,False);
  TobAEditer.AddChampSup('LIBELLEEMPLOI'     ,False);
  TobAEditer.AddChampSup('PSA_ETABLISSEMENT' ,False);
  TobAEditer.AddChampSup('PSA_DATEENTREE'    ,False);
  TobAEditer.AddChampSup('PSA_TRAVAILN1'     ,False);
  TobAEditer.AddChampSup('PSA_TRAVAILN2'     ,False);
  TobAEditer.AddChampSup('PSA_TRAVAILN3'     ,False);
  TobAEditer.AddChampSup('PSA_TRAVAILN4'     ,False);
  { Type de ligne }
  TobAEditer.AddChampSup('TYPECRITERE'       ,False);
  TobAEditer.AddChampSup('LESTAGE'           ,False);
  TobAEditer.AddChampSup('COMPETENCE'        ,False);
  TobAEditer.AddChampSup('LIBSTAGE'          ,False);
  TobAEditer.AddChampSup('LIBCOMPETENCE'     ,False);
  { Compétence }
  TobAEditer.AddChampSup('POS_COMPETENCE'    ,False);
  TobAEditer.AddChampSup('POS_LIBELLE'       ,False);
  TobAEditer.AddChampSup('POS_DEGREMAITRISE' ,False);
  TobAEditer.AddChampSup('POS_TABLELIBRECR1' ,False);
  TobAEditer.AddChampSup('POS_TABLELIBRECR2' ,False);
  TobAEditer.AddChampSup('POS_TABLELIBRECR3' ,False);
  TobAEditer.AddChampSup('PCO_TABLELIBRERH1' ,False);
  TobAEditer.AddChampSup('PCO_TABLELIBRERH2' ,False);
  TobAEditer.AddChampSup('PCO_TABLELIBRERH3' ,False);
  TobAEditer.AddChampSup('PCO_TABLELIBRERH4' ,False);
  TobAEditer.AddChampSup('PCO_TABLELIBRERH5' ,False);
  { Formation proposée }
  TobAEditer.AddChampSup('PST_DUREESTAGE'    ,False);
  TobAEditer.AddChampSup('PST_COUTBUDGETE'   ,False);
  TobAEditer.AddChampSup('PST_NATUREFORM'    ,False);
  TobAEditer.AddChampSup('PST_CENTREFORMGU'  ,False);
  TobAEditer.AddChampSup('PST_FORMATION1'    ,False);
  TobAEditer.AddChampSup('PST_FORMATION2'    ,False);
  TobAEditer.AddChampSup('PST_FORMATION3'    ,False);
  TobAEditer.AddChampSup('PST_FORMATION4'    ,False);
  TobAEditer.AddChampSup('PST_FORMATION5'    ,False);
  TobAEditer.AddChampSup('PST_FORMATION6'    ,False);
  TobAEditer.AddChampSup('PST_FORMATION7'    ,False);
  TobAEditer.AddChampSup('PST_FORMATION8'    ,False);
  { On récupère les noms des champs libres }
  AddTobChampsLibresRH(TobAEditer);
  TobCompetForm := Tob.Create('Compétences apportées par les formations',Nil,-1);
  TobCompetForm.LoadDetailFromSQL('SELECT POS_NATOBJSTAGE,POS_CODESTAGE,POS_ORDRE,POS_COMPETENCE,'+
    'POS_DEGREMAITRISE,POS_TABLELIBRECR1,POS_TABLELIBRECR2,POS_TABLELIBRECR3,'+
    'PST_LIBELLE, PST_DUREESTAGE,PST_COUTBUDGETE,PST_NATUREFORM,PST_CENTREFORMGU,PST_FORMATION1,PST_FORMATION2,'+
    'PST_FORMATION3,PST_FORMATION4,PST_FORMATION5,PST_FORMATION6,PST_FORMATION7,PST_FORMATION8 '+
    'FROM STAGEOBJECTIF '+
    'LEFT JOIN STAGE ON POS_CODESTAGE=PST_CODESTAGE '+
    'WHERE POS_NATOBJSTAGE="FOR" AND PST_MILLESIME="0000"');
{
  TobCompetEmpl := Tob.Create('Compétences liées aux emplois',Nil,-1);
  TobCompetEmpl.LoadDetailFromSQL('SELECT POS_NATOBJSTAGE,POS_CODESTAGE,POS_ORDRE,POS_COMPETENCE,'+
    'POS_DEGREMAITRISE,POS_TABLELIBRECR1,POS_TABLELIBRECR2,POS_TABLELIBRECR3,'+
    'PCO_TABLELIBRERH1,PCO_TABLELIBRERH2,PCO_TABLELIBRERH3,PCO_TABLELIBRERH4,PCO_TABLELIBRERH5 '+
    'FROM STAGEOBJECTIF '+
    'LEFT JOIN RHCOMPETENCES ON POS_COMPETENCE=PCO_COMPETENCE '+
    'WHERE POS_NATOBJSTAGE="EMP" '); }

  { Parcours des salariés }
  for i := 0 to TobSalarieCompetManq.Detail.Count -1 do
  begin
    TempTobSal := Tob.Create('Salarié - compétence - formation', nil, -1);
    TempTobSal.AddChampSupValeur('PSA_SALARIE',      TobSalarieCompetManq.Detail[i].Detail[0].GetValue('PSA_SALARIE'));
    TempTobSal.AddChampSupValeur('PSA_LIBELLE',      TobSalarieCompetManq.Detail[i].Detail[0].GetValue('PSA_LIBELLE'));
    TempTobSal.AddChampSupValeur('PSA_PRENOM',       TobSalarieCompetManq.Detail[i].Detail[0].GetValue('PSA_PRENOM'));
    TempTobSal.AddChampSupValeur('LIBELLEEMPLOI',    TobSalarieCompetManq.Detail[i].Detail[0].GetValue('LIBELLEEMPLOI'));
    TempTobSal.AddChampSupValeur('PSA_ETABLISSEMENT',TobSalarieCompetManq.Detail[i].Detail[0].GetValue('PSA_ETABLISSEMENT'));
    TempTobSal.AddChampSupValeur('PSA_DATEENTREE',   TobSalarieCompetManq.Detail[i].Detail[0].GetValue('PSA_DATEENTREE'));
    TempTobSal.AddChampSupValeur('PSA_TRAVAILN1',    TobSalarieCompetManq.Detail[i].Detail[0].GetValue('PSA_TRAVAILN1'));
    TempTobSal.AddChampSupValeur('PSA_TRAVAILN2',    TobSalarieCompetManq.Detail[i].Detail[0].GetValue('PSA_TRAVAILN2'));
    TempTobSal.AddChampSupValeur('PSA_TRAVAILN3',    TobSalarieCompetManq.Detail[i].Detail[0].GetValue('PSA_TRAVAILN3'));
    TempTobSal.AddChampSupValeur('PSA_TRAVAILN4',    TobSalarieCompetManq.Detail[i].Detail[0].GetValue('PSA_TRAVAILN4'));
    { Parcours des compétences manquantes }
    for j := 0 to TobSalarieCompetManq.Detail[i].Detail.Count -1 do
    begin
      TempTobSalCompet := TobSalarieCompetManq.Detail[i].Detail[j];
      { Ajout de la compétence salarié à la TobAEditer }
      TempTobLigneEdition := Tob.Create('Salarié - compétence - formation', TobAEditer, -1);
      TempTobLigneEdition.Dupliquer(TempTobSal, True, True);
      TempTobLigneEdition.AddChampSupValeur('TYPECRITERE'       ,'1');
      TempTobLigneEdition.AddChampSupValeur('LESTAGE'           ,'');
      TempTobLigneEdition.AddChampSupValeur('COMPETENCE'        ,TempTobSalCompet.GetValue('POS_COMPETENCE'));
      TempTobLigneEdition.AddChampSupValeur('LIBSTAGE'          ,'');
      TempTobLigneEdition.AddChampSupValeur('LIBCOMPETENCE'     ,TempTobSalCompet.GetValue('POS_LIBELLE'));
      TempTobLigneEdition.AddChampSupValeur('POS_COMPETENCE'    ,TempTobSalCompet.GetValue('POS_COMPETENCE'));
      TempTobLigneEdition.AddChampSupValeur('POS_LIBELLE'       ,TempTobSalCompet.GetValue('POS_LIBELLE'));
      TempTobLigneEdition.AddChampSupValeur('POS_DEGREMAITRISE' ,TempTobSalCompet.GetValue('POS_DEGREMAITRISE'));
      TempTobLigneEdition.AddChampSupValeur('POS_TABLELIBRECR1' ,TempTobSalCompet.GetValue('POS_TABLELIBRECR1'));
      TempTobLigneEdition.AddChampSupValeur('POS_TABLELIBRECR2' ,TempTobSalCompet.GetValue('POS_TABLELIBRECR2'));
      TempTobLigneEdition.AddChampSupValeur('POS_TABLELIBRECR3' ,TempTobSalCompet.GetValue('POS_TABLELIBRECR3'));
      TempTobLigneEdition.AddChampSupValeur('PCO_TABLELIBRERH1' ,RechDom('PGTABLELIBRERH1',TempTobSalCompet.GetValue('PCO_TABLELIBRERH1'),False));
      TempTobLigneEdition.AddChampSupValeur('PCO_TABLELIBRERH2' ,RechDom('PGTABLELIBRERH2',TempTobSalCompet.GetValue('PCO_TABLELIBRERH2'),False));
      TempTobLigneEdition.AddChampSupValeur('PCO_TABLELIBRERH3' ,RechDom('PGTABLELIBRERH3',TempTobSalCompet.GetValue('PCO_TABLELIBRERH3'),False));
      TempTobLigneEdition.AddChampSupValeur('PCO_TABLELIBRERH4' ,RechDom('PGTABLELIBRERH4',TempTobSalCompet.GetValue('PCO_TABLELIBRERH4'),False));
      TempTobLigneEdition.AddChampSupValeur('PCO_TABLELIBRERH5' ,RechDom('PGTABLELIBRERH5',TempTobSalCompet.GetValue('PCO_TABLELIBRERH5'),False));
      TempTobLigneEdition.AddChampSupValeur('PST_DUREESTAGE'    ,'');
      TempTobLigneEdition.AddChampSupValeur('PST_COUTBUDGETE'   ,'');
      TempTobLigneEdition.AddChampSupValeur('PST_NATUREFORM'    ,'');
      TempTobLigneEdition.AddChampSupValeur('PST_CENTREFORMGU'  ,'');
      TempTobLigneEdition.AddChampSupValeur('PST_FORMATION1'    ,'');
      TempTobLigneEdition.AddChampSupValeur('PST_FORMATION2'    ,'');
      TempTobLigneEdition.AddChampSupValeur('PST_FORMATION3'    ,'');
      TempTobLigneEdition.AddChampSupValeur('PST_FORMATION4'    ,'');
      TempTobLigneEdition.AddChampSupValeur('PST_FORMATION5'    ,'');
      TempTobLigneEdition.AddChampSupValeur('PST_FORMATION6'    ,'');
      TempTobLigneEdition.AddChampSupValeur('PST_FORMATION7'    ,'');
      TempTobLigneEdition.AddChampSupValeur('PST_FORMATION8'    ,'');
      { On récupère les noms des champs libres }
      AddTobChampsLibresRH(TempTobLigneEdition);
      { Parcours des formations correspondantes }
      TempTobFormation := TobCompetForm.FindFirst(['POS_COMPETENCE'], [TobSalarieCompetManq.Detail[i].Detail[j].GetString('POS_COMPETENCE')], False);
      While Assigned(TempTobFormation) do
      begin
        { Ajout de la formation à la TobAEditer }
        TempTobLigneEdition := Tob.Create('Salarié - compétence - formation', TobAEditer, -1);
        TempTobLigneEdition.Dupliquer(TempTobSal, True, True);
        TempTobLigneEdition.AddChampSupValeur('TYPECRITERE'     ,'2');
        TempTobLigneEdition.AddChampSupValeur('LESTAGE'         ,TempTobFormation.GetValue('POS_CODESTAGE'));
        TempTobLigneEdition.AddChampSupValeur('COMPETENCE'      ,TempTobSalCompet.GetValue('POS_COMPETENCE'));
        TempTobLigneEdition.AddChampSupValeur('LIBSTAGE'        ,TempTobFormation.GetValue('PST_LIBELLE'));
        TempTobLigneEdition.AddChampSupValeur('LIBCOMPETENCE'   ,TempTobSalCompet.GetValue('LIBCOMPETENCE'));
        TempTobLigneEdition.AddChampSupValeur('POS_COMPETENCE'    ,'');
        TempTobLigneEdition.AddChampSupValeur('POS_LIBELLE'       ,'');
        TempTobLigneEdition.AddChampSupValeur('POS_DEGREMAITRISE' ,'');
        TempTobLigneEdition.AddChampSupValeur('POS_TABLELIBRECR1' ,TempTobFormation.GetValue('POS_TABLELIBRECR1'));
        TempTobLigneEdition.AddChampSupValeur('POS_TABLELIBRECR2' ,TempTobFormation.GetValue('POS_TABLELIBRECR2'));
        TempTobLigneEdition.AddChampSupValeur('POS_TABLELIBRECR3' ,TempTobFormation.GetValue('POS_TABLELIBRECR3'));
        TempTobLigneEdition.AddChampSupValeur('PCO_TABLELIBRERH1' ,'');
        TempTobLigneEdition.AddChampSupValeur('PCO_TABLELIBRERH2' ,'');
        TempTobLigneEdition.AddChampSupValeur('PCO_TABLELIBRERH3' ,'');
        TempTobLigneEdition.AddChampSupValeur('PCO_TABLELIBRERH4' ,'');
        TempTobLigneEdition.AddChampSupValeur('PCO_TABLELIBRERH5' ,'');
        TempTobLigneEdition.AddChampSupValeur('PST_DUREESTAGE'  ,TempTobFormation.GetValue('PST_DUREESTAGE'));
        TempTobLigneEdition.AddChampSupValeur('PST_COUTBUDGETE' ,TempTobFormation.GetValue('PST_COUTBUDGETE'));
        TempTobLigneEdition.AddChampSupValeur('PST_NATUREFORM'  ,TempTobFormation.GetValue('PST_NATUREFORM'));
        TempTobLigneEdition.AddChampSupValeur('PST_CENTREFORMGU',TempTobFormation.GetValue('PST_CENTREFORMGU'));
        TempTobLigneEdition.AddChampSupValeur('PST_FORMATION1'  ,TempTobFormation.GetValue('PST_FORMATION1'));
        TempTobLigneEdition.AddChampSupValeur('PST_FORMATION2'  ,TempTobFormation.GetValue('PST_FORMATION2'));
        TempTobLigneEdition.AddChampSupValeur('PST_FORMATION3'  ,TempTobFormation.GetValue('PST_FORMATION3'));
        TempTobLigneEdition.AddChampSupValeur('PST_FORMATION4'  ,TempTobFormation.GetValue('PST_FORMATION4'));
        TempTobLigneEdition.AddChampSupValeur('PST_FORMATION5'  ,TempTobFormation.GetValue('PST_FORMATION5'));
        TempTobLigneEdition.AddChampSupValeur('PST_FORMATION6'  ,TempTobFormation.GetValue('PST_FORMATION6'));
        TempTobLigneEdition.AddChampSupValeur('PST_FORMATION7'  ,TempTobFormation.GetValue('PST_FORMATION7'));
        TempTobLigneEdition.AddChampSupValeur('PST_FORMATION8'  ,TempTobFormation.GetValue('PST_FORMATION8'));
        TempTobFormation := TobCompetForm.FindNext(['POS_COMPETENCE'], [TobSalarieCompetManq.Detail[i].Detail[j].GetString('POS_COMPETENCE')], False);
        { On récupère les noms des champs libres }
        AddTobChampsLibresRH(TempTobLigneEdition);
      end;
    end;
    FreeAndNil(TempTobSal);
  end;
  FreeAndNil(TobCompetForm);
  FreeAndNil(TobSalarieCompetManq);
  TobAEditer.Detail.Sort('PSA_SALARIE;TYPECRITERE;LESTAGE');
  TFQRS1(Ecran).LaTob:= TobAEditer;

end;

procedure TOF_PGEDITCPTCESACQUER.OnClose;
begin
  inherited;
  If Assigned(TobAEditer) then FreeAndNil(TobAEditer);
end;

initialization
  registerclasses([TOF_PGEDITCPTCESACQUER]);
end.
