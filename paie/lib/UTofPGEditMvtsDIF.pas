{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 25/04/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGEDITCOMPTEURSDIF ()
Mots clefs ... : TOF;PGEDITCOMPTEURSDIF
*****************************************************************
PT1  | 25/02/2008 | V_803 | FL | Edition dans la fenêtre mère
}
Unit UTofPGEditMvtsDIF ;

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
     HSysMenu,
     ed_tools;

Type
  TOF_PGEDITMVTSDIF = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnUpdate                 ; override ; //PT1
    private
    TobEditDIF : Tob;
    //procedure LancerDIF(Sender : TObject); //PT1
    procedure ExitEdit(Sender: TObject);
  end ;

Implementation

procedure TOF_PGEDITMVTSDIF.OnUpdate;
var Q : TQuery;
    Pages : TPageControl;
    Where,WhereAbs,WhereRub : String;
    TE,TobMvtDIF,TobFormation : Tob;
    m,a : Integer;
    DD,DF,DDFor,DFFor,DDRub,DFRub,DatePaieF : TDateTime;
    mois,annee : String;
    TypeConge,Mouvement,Rubrique : String;
    TypePlan,WhereFormation,Salarie : String;
begin
  Mois := GetControlText('CBXMOIS');
  Annee := GetControlText('CBXANNEE');
  If Annee <> '' then
  begin
    Q := OpenSQL('SELECT PEX_ANNEEREFER FROM EXERSOCIAL WHERE PEX_EXERCICE="'+Annee+'"',True);
    If Not Q.Eof then Annee := Q.FindField('PEX_ANNEEREFER').AsString;
    Ferme(Q);
    If Mois = '' then
    begin
        DD := encodedate(StrToInt(Annee),01, 01);
        DF := encodedate(StrToInt(Annee),12, 31);
    end
    else
    begin
        DD := encodedate(StrToInt(Annee),StrToInt(Mois), 01);
        DF := FinDeMois(DD);
    end;
  end
  else
  begin
        DD := IDate1900;
        DF := encodedate(2099,12, 31);
  end;
  Pages := TPageControl(GetControl('PAGES'));
  Where := RecupWhereCritere(Pages);
  WhereAbs := ConvertPrefixe(Where,'PSA','PCN');
  If WhereAbs <> '' then WhereAbs := WhereAbs + 'AND '+
             'PCN_DATEFINABS>="'+UsDateTime(DD)+'" AND '+
             'PCN_DATEFINABS<="'+UsDateTime(DF)+'"'
  else WhereAbs := 'WHERE '+
             'PCN_DATEFINABS>="'+UsDateTime(DD)+'" AND '+
             'PCN_DATEFINABS<="'+UsDateTime(DF)+'"';
  WhereRub := ConvertPrefixe(Where,'PSA','PSD');
  If WhereRub <> '' then WhereRUb := WhereRub   + 'AND '+
            '(PSD_ORIGINEMVT="MHE" OR PSD_ORIGINEMVT="MLB") AND '+
            'PSD_DATEFIN>="'+UsDateTime(DD)+'" AND '+
            'PSD_DATEFIN<="'+UsDateTime(DF)+'" AND '+
            'PSD_ORDRE>=200'
  else WhereRub := 'WHERE '+
            '(PSD_ORIGINEMVT="MHE" OR PSD_ORIGINEMVT="MLB") AND '+
            'PSD_DATEFIN>="'+UsDateTime(DD)+'" AND '+
            'PSD_DATEFIN<="'+UsDateTime(DF)+'" AND '+
            'PSD_ORDRE>=200';
  TobEditDIF := Tob.Create('AffichageMvt',Nil,-1);
  
  Q := OpenSQL('SELECT * FROM PARAMFORMABS WHERE ##PPF_PREDEFINI##',True);
  TobMvtDIF := Tob.Create('Lesmvts',Nil,-1);
  TobMvtDIF.LoadDetailDB('Lesmvts','','',Q,False);
  Ferme(Q);
  
  For m := 0 TO TobMvtDIF.Detail.Count - 1 do
  begin
       If TobMvtDIF.Detail[m].GetValue('PPF_ALIMABS') = 'X' then Mouvement := 'ABSENCE'
       else If TobMvtDIF.Detail[m].GetValue('PPF_ALIMRUB') = 'X' then Mouvement := 'RUBRIQUE'
       else Mouvement := '';
       TypePlan := TobMvtDIF.Detail[m].GetValue('PPF_TYPEPLANPREV');
       WhereFormation := '';
       If TobMvtDIF.Detail[m].GetValue('PPF_CHEURETRAV') = 'X' then WhereFormation := ' AND PFO_HTPSTRAV>0';
       If TobMvtDIF.Detail[m].GetValue('PPF_CHEURENONTRAV') = 'X' then WhereFormation := WhereFormation + ' AND PFO_HTPSNONTRAV>0';
       
       Q := OpenSQL('SELECT PFO_SALARIE,PFO_TYPEPLANPREV,PFO_CODESTAGE,PFO_DATEDEBUT,PFO_DATEFIN,PFO_NBREHEURE,PFO_DATEPAIE '+
       'FROM FORMATIONS WHERE PFO_DATEFIN>="'+UsDateTime(DD)+'" AND PFO_DATEFIN <="'+UsDateTime(DF)+'" '+
       'AND PFO_EFFECTUE="X" AND PFO_TYPEPLANPREV="'+TypePlan+'"'+WhereFormation,True);
       TobFormation := Tob.Create('LesFormations',Nil,-1);
       TobFormation.LoadDetailDB('LesFormations','','',Q,False);
       Ferme(Q);
       
       For a := 0 to TobFormation.detail.Count - 1 do
       begin
            TE := Tob.Create('filleEdition',TobEditDIF,-1);
            Salarie := TobFormation.Detail[a].GetValue('PFO_SALARIE');
            TE.AddChampSupValeur('SALARIE',Salarie);
            TE.AddChampSupValeur('LIBELLE',RechDom('PGSALARIE',Salarie,False));
            TE.AddChampSupValeur('LIBMOUVEMENT',RechDom('PGSTAGEFORM',TobFormation.Detail[a].GetValue('PFO_CODESTAGE'),False));
            TE.AddChampSupValeur('TYPEFOR',RechDom('PGTYPEPLANPREV',TobFormation.Detail[a].GetValue('PFO_TYPEPLANPREV'),False));
            DDFor := TobFormation.Detail[a].GetValue('PFO_DATEDEBUT');
            DFFor := TobFormation.Detail[a].GetValue('PFO_DATEFIN');
            TE.AddChampSupValeur('DATEDEB',DDFor);
            TE.AddChampSupValeur('DATEFIN',DFFor);
            TE.AddChampSupValeur('HEURES',TobFormation.Detail[a].GetValue('PFO_NBREHEURE'));
            DatePaieF := TobFormation.Detail[a].GetValue('PFO_DATEPAIE');
            If Mouvement = 'ABSENCE' then
            begin
                 TypeConge := TobMvtDIF.Detail[m].GetValue('PPF_TYPECONGE');
                 WhereAbs := 'WHERE PCN_SALARIE="'+Salarie+'" AND '+
                 'PCN_DATEFINABS>="'+UsDateTime(DDFor)+'" AND '+
                 'PCN_DATEFINABS<="'+UsDateTime(DFFor)+'"';
                 Q := OpenSQL('SELECT * FROM ABSENCESALARIE '+WhereAbs+' AND PCN_TYPECONGE="'+TypeConge+'"',True);
                 If Not Q.Eof then
                 begin
                      TE.AddChampSupValeur('MOUVEMENT','Absence');
                      TE.AddChampSupValeur('DATEINTEGR',Q.FindField('PCN_DATEPAIEMENT').AsDateTime);
                      If Q.FindField('PCN_CODETAPE').AsString = 'P' then TE.AddChampSupValeur('PAYE','X')
                      else TE.AddChampSupValeur('PAYE','-');
                 end
                 else
                 begin
                      TE.AddChampSupValeur('MOUVEMENT','');
                      TE.AddChampSupValeur('DATEINTEGR',IDate1900);
                      TE.AddChampSupValeur('PAYE','-');
                 end;
                 Ferme(Q);
            end
            else
            if Mouvement = 'RUBRIQUE' then
            begin
                 Rubrique := TobMvtDIF.Detail[m].GetValue('PPF_RUBRIQUE');
                 DDRub := DebutDeMois(DatePaieF);
                 DFRub := FinDemois(DatePaieF);
                 WhereRub := 'WHERE PSD_SALARIE="'+Salarie+'" AND '+
                             '(PSD_ORIGINEMVT="MHE" OR PSD_ORIGINEMVT="MLB") AND '+
                             'PSD_DATEFIN>="'+UsDateTime(DDRub)+'" AND '+
                             'PSD_DATEFIN<="'+UsDateTime(DFRub)+'" AND '+
                             'PSD_ORDRE>=200';
                 Q := OpenSQL('SELECT * FROM HISTOSAISRUB '+WhereRub+' AND PSD_RUBRIQUE="'+Rubrique+'" AND PSD_ORIGINEMVT="MLB"',True);
                 If Not Q.Eof then
                 begin
                      TE.AddChampSupValeur('MOUVEMENT','Rubrique');
                      TE.AddChampSupValeur('DATEINTEGR',Q.FindField('PSD_DATEINTEGRAT').AsDateTime);
                      If Q.FindField('PSD_DATEINTEGRAT').AsDateTime > (IDate1900) then TE.AddChampSupValeur('PAYE','X')
                      else TE.AddChampSupValeur('PAYE','-');
                 end
                 else
                 begin
                      TE.AddChampSupValeur('MOUVEMENT','');
                      TE.AddChampSupValeur('DATEINTEGR',IDate1900);
                      TE.AddChampSupValeur('PAYE','-');
                 end;
                 Ferme(Q);
            end;
       end;
       TobFormation.free;
  end;
  TobMvtDIF.Free;

  //PT1  
  TFQRS1(Ecran).LaTob := TobEditDIF;
  //If GetCheckBoxState('FLISTE') = CbChecked then LanceEtatTOB('E','PDI','PMV',TobEditDIF,True,True,False,Pages,' ','',False)
  //else LanceEtatTOB('E','PDI','PMV',TobEditDIF,True,False,False,Pages,' ','',False);
  //TobEditDIF.Free;
end ;

procedure TOF_PGEDITMVTSDIF.OnArgument (S : String ) ;
var //BVal : TToolBarButton97;
    Min,Max : String;
    Defaut: THEdit;
    MoisE, AnneeE, ComboExer : String;
    DebExer, FinExer : TDateTime;
begin
  Inherited ;
  //PT1 - Début
  //BVal := TToolBarButton97(GetControl('BLANCEDIF')); 
  //If BVal <> Nil then BVal.OnClick := LancerDIF;
  //BVal := TToolBarButton97(GetControl('BValider')); //FC le 08/02/07 pour que le F9 fonctionne
  //If BVal <> Nil then BVal.OnClick := LancerDIF;
  //PT1 - Fin
  SetControlText('DOSSIER',GetParamSoc ('SO_LIBELLE'));
  RecupMinMaxTablette('PG','SALARIES','PSA_SALARIE',Min,Max);
  SetControlText('PSA_SALARIE',Min);
  SetControlText('PSA_SALARIE_',Max);
  RecupMinMaxTablette('PG', 'ETABLISS', 'ET_ETABLISSEMENT', Min, Max);
  SetControlText('PSA_ETABLISSEMENT', Min);
  SetControlText('PSA_ETABLISSEMENT_', Max);
  Defaut:=ThEdit(getcontrol('PSA_SALARIE'));
  If Defaut<>nil then Defaut.OnExit:=ExitEdit;
  Defaut:=ThEdit(getcontrol('PSA_SALARIE_'));
  If Defaut<>nil then Defaut.OnExit:=ExitEdit;
  if RendExerSocialEnCours(MoisE, AnneeE, ComboExer, DebExer, FinExer) = TRUE then
  begin
    SetControlText('CBXMOIS',MoisE);
    SetControlText('CBXANNEE',ComboExer);
  end;
end ;

procedure TOF_PGEDITMVTSDIF.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;


Initialization
  registerclasses ( [ TOF_PGEDITMVTSDIF ] ) ;
end.

