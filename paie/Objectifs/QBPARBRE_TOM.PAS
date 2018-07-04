{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 15/04/2005
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : QBPARBRE (QBPARBRE)
Mots clefs ... : TOM;QBPARBRE
*****************************************************************}
Unit QBPARBRE_TOM ;                                                   

Interface

Uses HEnt1,UTOM,
     {$IFNDEF EAGLCLIENT} db,
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
     {$else} utob, {$ENDIF}
     UTableFiltre;

Type
  TOM_QBPARBRE = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
 private
     TF: TTableFiltre;
     okNew:boolean;
     evolprctSvg,saivalSvg,evolvalSvg:hString;
     EvolQtePrctSvg,SaisieQteValSvg,EvolQteValSvg:hString;
     EvolCAPrct2Svg,SaisieCAVal2Svg,EvolCAVal2Svg:hString;
     EvolCAPrct3Svg,SaisieCAVal3Svg,EvolCAVal3Svg:hString;
     EvolCAPrct4Svg,SaisieCAVal4Svg,EvolCAVal4Svg:hString;
     EvolCAPrct5Svg,SaisieCAVal5Svg,EvolCAVal5Svg:hString;
     EvolCAPrct6Svg,SaisieCAVal6Svg,EvolCAVal6Svg:hString;
     { EVI / ORLI : Onglets Prevision }
     //SaisieCAValPrevSvg,EvolCAValPrevSvg:hString;
     //EvolCAPrctPrev,EvolQtePrctPrev,SaisieQteValPrevSvg,EvolQteValPrevSvg:hString;
     CoeffRetenuSvg,NewCoeff:double;
     OkModifC,OkModifS:boolean;
     OkModifSaisieQteValS,OkModifEvolQteS:boolean;
     OkModifSaisieCAVal2S,OkModifEvolCA2S:boolean;
     OkModifSaisieCAVal3S,OkModifEvolCA3S:boolean;
     OkModifSaisieCAVal4S,OkModifEvolCA4S:boolean;
     OkModifSaisieCAVal5S,OkModifEvolCA5S:boolean;
     OkModifSaisieCAVal6S,OkModifEvolCA6S:boolean;
     { EVI / ORLI : Onglets Prevision }
     //OkModifSaisieCAValPrevS,OkModifEvolCAPrevS:boolean;
     //OkModifSaisieQteValPrevS,OkModifEvolQtePrevS:boolean;
     OkModifCoeffRetenu:boolean;
     nivMaxSession,NbIntervalle:integer;
     OkInitCoeff,OkInitPrev,OkInitObj,OkInitNivTaille:boolean;
     procedure AffichageColonneListe;
     procedure Update(const EdtSaisi,EdtCorrectif,
               ChpSaisi,ChpCorrectif,ChpCorrectifPrct,SaisiSvg,CorrectifSvg:hString;
               var OkModifSaisi,OkModifCorrectif:boolean);
     procedure AfterUpdate(okQte:boolean;const ChpCorrectifPrct,EdtSaisie,EdtCorrectif,
               CorrectifPrctSvg,SaisieSvg,CorrectifSvg,ChpNewVal,ChpVal:hString;
               OkModifSaisi,OkModifCorrectif:boolean;
               var codeSql,codeSqlNewValeurNoeud:hString);
     procedure CodeSqlMAJPrct(const codeSession,NoeudPere:hString;
               var CodeSql:hString);
     procedure Load(const edtSaisi,edtCorrectif,ChpSaisi,ChpCorrectif,ChpCorrectifPrct:hString;
               var EvolPrctSvg,SaisieValSvg,CorrectifSvg:hString);
    end ;

var codeaxeT,numnoeudT,numnoeudpereT,niveauT,numNoeudDup:hString;
    ValAxeNiv1T,ValAxeNiv2T,ValAxeNiv3T,ValAxeNiv4T:hString;
    ValAxeNiv5T,ValAxeNiv6T,ValAxeNiv7T,ValAxeNiv8T,ValAxeNiv9T,ValAxeNiv10T:hString;
    OkModifValBloque:boolean;
    OkRepriseHisto,OkDelete,OkDuplication:boolean;
    TabValues : array[1..7] of String;

Implementation

Uses StdCtrls,Classes,Graphics,Sysutils,HCtrls,CstCommun,
     HMsgBox,SaisieList,BPBasic,BPFctArbre,BPFctSession,UUtil,BPUtil;

procedure TOM_QBPARBRE.OnNewRecord ;
begin
  Inherited ;
  okNew:=true;
  SetField('QBR_CODEAXE',codeaxeT);
  SetField('QBR_NUMNOEUD',numnoeudT);
  SetField('QBR_NUMNOEUDPERE',numnoeudpereT);
  SetField('QBR_NIVEAU',niveauT);
  SetField('QBR_VALAXENIV1',ValAxeNiv1T);
  SetField('QBR_VALAXENIV2',ValAxeNiv2T);
  SetField('QBR_VALAXENIV3',ValAxeNiv3T);
  SetField('QBR_VALAXENIV4',ValAxeNiv4T);
  SetField('QBR_VALAXENIV5',ValAxeNiv5T);
  SetField('QBR_VALAXENIV6',ValAxeNiv6T);
  SetField('QBR_VALAXENIV7',ValAxeNiv7T);
  SetField('QBR_VALAXENIV8',ValAxeNiv8T);
  SetField('QBR_VALAXENIV9',ValAxeNiv9T);
  SetField('QBR_VALAXENIV10',ValAxeNiv10T);

  { EVI / ToDo : gestion de la devise pour les axes manuels }
  SetField('QBR_DEVISE','"' + V_PGI.DevisePivot + '"');

  if not OkRepriseHisto
   then
    begin
     //CA1
     SetField('QBR_REF1',0);
     SetField('QBR_REFPRCT1',0);
//     SetField('QBR_OP1',0);
//     SetField('QBR_OPPRCT1',0);
//     SetField('QBR_EVOLVAL1',0);
//     SetField('QBR_EVOLPRCT1',0);
//     SetField('QBR_SAISI1',0);
     //CA2
     SetField('QBR_REF2',0);
     SetField('QBR_REFPRCT2',0);
//     SetField('QBR_OP2',0);
//     SetField('QBR_OPPRCT2',0);
//     SetField('QBR_EVOLVAL2',0);
//     SetField('QBR_EVOLPRCT2',0);
//     SetField('QBR_SAISI2',0);
     //CA3
     SetField('QBR_REF3',0);
     SetField('QBR_REFPRCT3',0);
//     SetField('QBR_OP3',0);
//     SetField('QBR_OPPRCT3',0);
//     SetField('QBR_EVOLVAL3',0);
//     SetField('QBR_EVOLPRCT3',0);
//     SetField('QBR_SAISI3',0);
     //CA4
     SetField('QBR_REF4',0);
     SetField('QBR_REFPRCT4',0);
//     SetField('QBR_OP4',0);
//     SetField('QBR_OPPRCT4',0);
//     SetField('QBR_EVOLVAL4',0);
//     SetField('QBR_EVOLPRCT4',0);
//     SetField('QBR_SAISI4',0);
     //CA5
     SetField('QBR_REF5',0);
     SetField('QBR_REFPRCT5',0);
//     SetField('QBR_OP5',0);
//     SetField('QBR_OPPRCT5',0);
//     SetField('QBR_EVOLVAL5',0);
//     SetField('QBR_EVOLPRCT5',0);
  //   SetField('QBR_SAISI5',0);
     //CA6
     SetField('QBR_REF6',0);
     SetField('QBR_REFPRCT6',0);
//     SetField('QBR_OP6',0);
//     SetField('QBR_OPPRCT6',0);
//     SetField('QBR_EVOLVAL6',0);
//     SetField('QBR_EVOLPRCT6',0);
//     SetField('QBR_SAISI6',0);
     //QTE
     SetField('QBR_QTEREF',0);
     SetField('QBR_QTEREFPRCT',0);
//     SetField('QBR_QTEC',0);
//     SetField('QBR_QTECPRCT',0);
//     SetField('QBR_EVOLVALQTE',0);
//     SetField('QBR_EVOLPRCTQTE',0);
//     SetField('QBR_SAISIQTE',0);
    end;
  THEdit(GetControl('QBR_VALEURAXE')).Enabled:=true;
//  THEdit(GetControl('QBR_SAISIVAL')).Enabled:=true;
//  THEdit(GetControl('QBR_CORRECTIFVAL')).Enabled:=true;
//  THEdit(GetControl('QBR_CORRECTIFPRCT')).Enabled:=true;

  if codeaxeT='DELAI'
  then
   begin
    THEdit(GetControl('QBR_VALEURAXE')).visible:=false;
    THEdit(GetControl('QBR_DATEDELAI')).visible:=true;
   end;

  if ((ContextBP=3) and (codeaxeT='011')) then THCheckBox(GetControl('CB_NEWVALUE')).Visible:=true;
end ;



procedure TOM_QBPARBRE.OnDeleteRecord ;
var codeSession,noeudpere,noeud,codeSql:hString;
    valeurCA1,valeurQte,valeurCA2,valeurCA3,valeurCA4,valeurCA5,valeurCA6:double;
    valeurCA1Ref,valeurQteRef,valeurCA2Ref,valeurCA3Ref,valeurCA4Ref,valeurCA5Ref,valeurCA6Ref:double;
    TotalQte,totalniv,TotalCA2,TotalCA3,TotalCA4,TotalCA5,TotalCA6,TotalPrevu,TotalRet,TotalCaRetenu:double;
    TotalQteRef,totalCA1Ref,TotalCA2Ref,TotalCA3Ref:double;
    TotalCA4Ref,TotalCA5Ref,TotalCA6Ref:double;
    TotalNQte,TotalNCA1,TotalNCA2,TotalNCA3,TotalNCA4:double;
    TotalNCA5,TotalNCA6,TotalPrevuN,TotalRetN:double;
    niveau:hString;
begin
 Inherited ;
 OkDelete:=true;

 codeSession:=GetField('QBR_CODESESSION');
 noeudPere:=GetField('QBR_NUMNOEUDPERE');
 niveau:=GetField('QBR_NIVEAU');

 valeurQte:=GetField('QBR_QTEC');
 valeurCA1:=GetField('QBR_OP1');
 valeurCA2:=GetField('QBR_OP2');
 valeurCA3:=GetField('QBR_OP3');
 valeurCA4:=GetField('QBR_OP4');
 valeurCA5:=GetField('QBR_OP5');
 valeurCA6:=GetField('QBR_OP6');

 valeurQteRef:=GetField('QBR_QTEREF');
 valeurCA1Ref:=GetField('QBR_REF1');
 valeurCA2Ref:=GetField('QBR_REF2');
 valeurCA3Ref:=GetField('QBR_REF3');
 valeurCA4Ref:=GetField('QBR_REF4');
 valeurCA5Ref:=GetField('QBR_REF5');
 valeurCA6Ref:=GetField('QBR_REF6');

 noeud:=GetField('QBR_NUMNOEUD');


 //MAj des valeurs à 0
 codeSql:=' QBR_OP1="0",QBR_OP2="0",QBR_OP3="0",'+
          'QBR_OP4="0",QBR_OP5="0",QBR_OP6="0",QBR_QTEC="0" ';

 //total valeur niveau
 //on recupère les valeurs Histo avant modification
 ArbreTotalNivRef(codeSession,noeudPere,
                  TotalQteRef,totalCA1Ref,TotalCA2Ref,TotalCA3Ref,
                  TotalCA4Ref,TotalCA5Ref,TotalCA6Ref);
 //on recupère les valeurs courante avant modification
 ArbreTotalNiv(codeSession,noeudPere,
               TotalQte,totalniv,TotalCA2,TotalCA3,TotalCA4,TotalCA5,TotalCA6,
               TotalPrevu,TotalRet,TotalCaRetenu);

 //maj du noeud
 MExecuteSql('UPDATE QBPARBRE SET '+codeSql+
          ' WHERE QBR_CODESESSION="'+codeSession+
          '" AND QBR_NUMNOEUD="'+Noeud+'"',
          'QBPARBRE_TOM (OnAfterUpdateRecord).');

 ArbreTotalNiv(codeSession,noeudPere,
            TotalNQte,TotalNCA1,TotalNCA2,TotalNCA3,TotalNCA4,
            TotalNCA5,TotalNCA6,TotalPrevuN,TotalRetN,TotalCaRetenu);
 MiseAJourValeurBloque(codeSession,StrToInt(niveau));
 MiseAjourNiv(true,codeSession,Noeud,noeudPere,niveau,nivMaxSession,
           TotalNQte,TotalNCA1,TotalNCA2,TotalNCA3,
           TotalNCA4,TotalNCA5,TotalNCA6,
           TotalNQte-TotalQteRef,TotalNCA1-totalCA1Ref,
           TotalNCA2-totalCA2Ref,TotalNCA3-totalCA3Ref,
           TotalNCA4-totalCA4Ref,TotalNCA5-totalCA5Ref,
           TotalNCA6-totalCA6Ref,
           STRFPOINT(NewCoeff),OkModifCoeffRetenu,OkInitPrev,OkInitNivTaille);


 TotalNiv:=totalniv-valeurCA1;
 TotalQte:=totalQte-valeurQte;
 TotalCA2:=totalCA2-valeurCA2;
 TotalCA3:=totalCA3-valeurCA3;
 TotalCA4:=totalCA4-valeurCA4;
 TotalCA5:=totalCA5-valeurCA5;
 TotalCA6:=totalCA6-valeurCA6;

 //mise à jour des %
 codeSql:='';
 if TotalNiv>0
  then codeSql:='QBR_OPPRCT1=((QBR_OP1/'+strFPoint4(totalniv)+')*100) ';
 if TotalQte>0
  then
   begin
    if codeSql<>''
     then codeSql:=codeSql+',';
    codeSql:=codeSql+'QBR_QTECPRCT=((QBR_QTEC/'+strFPoint4(TotalQte)+')*100) ';
   end;
 if TotalCA2>0
  then
   begin
    if codeSql<>''
     then codeSql:=codeSql+',';
     codeSql:=codeSql+'QBR_OPPRCT2=((QBR_OP2/'+strFPoint4(totalCA2)+')*100) ';
   end;
 if TotalCA3>0
  then
   begin
    if codeSql<>''
     then codeSql:=codeSql+',';
     codeSql:=codeSql+'QBR_OPPRCT3=((QBR_OP3/'+strFPoint4(totalCA3)+')*100) ';
   end;
 if TotalCA4>0
  then
   begin
    if codeSql<>''
     then codeSql:=codeSql+',';
     codeSql:=codeSql+'QBR_OPPRCT4=((QBR_OP4/'+strFPoint4(totalCA4)+')*100) ';
   end;
 if TotalCA5>0
  then
   begin
    if codeSql<>''
     then codeSql:=codeSql+',';
     codeSql:=codeSql+'QBR_OPPRCT5=((QBR_OP5/'+strFPoint4(totalCA5)+')*100) ';
   end;
 if TotalCA6>0
  then
   begin
    if codeSql<>''
     then codeSql:=codeSql+',';
     codeSql:=codeSql+'QBR_OPPRCT6=((QBR_OP6/'+strFPoint4(totalCA6)+')*100) ';
   end;

 if codeSql<>''
  then MExecuteSql('UPDATE QBPARBRE SET '+
                   codeSql+' WHERE QBR_CODESESSION="'+codeSession+
                   '" AND QBR_NUMNOEUDPERE="'+noeudPere+
                   '" AND QBR_NUMNOEUD<>"'+noeud+'"',
                   'QBPARBRE_TOM (OnDeleteRecord).');

  
 ArbreTotalNivHisto(codeSession,noeudPere,
                   TotalQteRef,totalCA1Ref,TotalCA2Ref,TotalCA3Ref,
                   TotalCA4Ref,TotalCA5Ref,TotalCA6Ref);
                   
 TotalCA1Ref:=totalCA1Ref-valeurCA1Ref;
 TotalQteRef:=totalQteRef-valeurQteRef;
 TotalCA2Ref:=totalCA2Ref-valeurCA2Ref;
 TotalCA3Ref:=totalCA3Ref-valeurCA3Ref;
 TotalCA4Ref:=totalCA4Ref-valeurCA4Ref;
 TotalCA5Ref:=totalCA5Ref-valeurCA5Ref;
 TotalCA6Ref:=totalCA6Ref-valeurCA6Ref;

 //mise de la valeur de reference du niveau superieur
 MExecuteSql('UPDATE QBPARBRE SET '+
            'QBR_REF1='+strFPoint4(totalCA1Ref)+' '+
            ',QBR_QTEREF='+strFPoint4(totalQteRef)+' '+
            ',QBR_REF2='+strFPoint4(totalCA2Ref)+' '+
            ',QBR_REF3='+strFPoint4(totalCA3Ref)+' '+
            ',QBR_REF4='+strFPoint4(totalCA4Ref)+' '+
            ',QBR_REF5='+strFPoint4(totalCA5Ref)+' '+
            ',QBR_REF6='+strFPoint4(totalCA6Ref)+' '+
            ' WHERE QBR_CODESESSION="'+codeSession+
            '" AND QBR_NUMNOEUD="'+noeudPere+'"',
            'QBPARBRE_TOM (OnDeleteRecord).');

  //suppression des valeurs de QBPARBREDETAIL
   MExecuteSql('DELETE FROM QBPARBREDETAIL WHERE '+
            'QBH_CODESESSION="'+codeSession+
            '" AND QBH_NUMNOEUDREF="'+noeud+'"',
            'QBPARBRE_TOM (OnDeleteRecord).');
end ;

procedure TOM_QBPARBRE.Update(const EdtSaisi,EdtCorrectif,
          ChpSaisi,ChpCorrectif,ChpCorrectifPrct,SaisiSvg,CorrectifSvg:hString;
          var OkModifSaisi,OkModifCorrectif:boolean);
var saivalS,correctifvalS:hString;
    saival,correctifval:double;
begin
 saivalS:=GetControlText(EdtSaisi);
 correctifvalS:=GetControlText(EdtCorrectif);
 OkModifSaisi:=false;
 OkModifCorrectif:=false;
 saival:=0;
 if (saivalS<>'')
  then
   begin
    saival:=Valeur(saivalS);
    OkModifSaisi:=true;
    SetControlText(EdtCorrectif,'');
    SetField(ChpCorrectifPrct,0);
   end
  else
   begin
    if (saivalS<>SaisiSvg) and (SaisiSvg<>'0')
     then
      begin
       OkModifSaisi:=true;
      end;
   end;
 correctifval:=0;
 if (correctifvalS<>'') or (correctifvalS<>CorrectifSvg)
  then
   begin
    correctifval:=Valeur(correctifvalS);
    OkModifCorrectif:=true;
   end;

 if GetControlText(EdtCorrectif)='' then correctifval:=0;

 SetField(ChpSaisi,saival);
 SetField(ChpCorrectif,correctifval);
end;

procedure TOM_QBPARBRE.OnUpdateRecord ;
begin
 Inherited ;
 //CA1
 Update('EDTSAISI1','EDTEVOLVAL1','QBR_SAISI1','QBR_EVOLVAL1',
        'QBR_EVOLPRCT1',
        saivalSvg,evolvalSvg,OkModifS,OkModifC);     
 //Qte
 Update('EDTSAISIQTE','EDTEVOLVALQTE','QBR_SAISIQTE','QBR_EVOLQTE',
        'QBR_EVOLQTEPRCT',
        SaisieQteValSvg,EvolQteValSvg,OkModifSaisieQteValS,OkModifEvolQteS);
 //CA2
 Update('EDTSAISI2','EDTEVOLVAL2','QBR_SAISI2','QBR_EVOLVAL2',
        'QBR_EVOLPRCT2',
        SaisieCAVal2Svg,EvolCAVal2Svg,OkModifSaisieCAVal2S,OkModifEvolCA2S);

 //CA3
 Update('EDTSAISI3','EDTEVOLVAL3','QBR_SAISI3','QBR_EVOLVAL3',
        'QBR_EVOLPRCT3',
        SaisieCAVal3Svg,EvolCAVal3Svg,OkModifSaisieCAVal3S,OkModifEvolCA3S);

 //CA4
 Update('EDTSAISI4','EDTEVOLVAL4','QBR_SAISI4','QBR_EVOLVAL4',
        'QBR_EVOLPRCT4',
        SaisieCAVal4Svg,EvolCAVal4Svg,OkModifSaisieCAVal4S,OkModifEvolCA4S);

 //CA5
 Update('EDTSAISI5','EDTEVOLVAL5','QBR_SAISI5','QBR_EVOLVAL5',
        'QBR_EVOLPRCT5',
        SaisieCAVal5Svg,EvolCAVal5Svg,OkModifSaisieCAVal5S,OkModifEvolCA5S);

 //CA6
 Update('EDTSAISI6','EDTEVOLVAL6','QBR_SAISI6','QBR_EVOLVAL6',
        'QBR_EVOLPRCT6',
        SaisieCAVal6Svg,EvolCAVal6Svg,OkModifSaisieCAVal6S,OkModifEvolCA6S);

 { EVI / ORLI : Onglets CA & QTE Prevision }
 {//CA prev
 Update('EDTSAISIPREVCA','EDTEVOLVALPREVCA','QBR_SAISIPREVCA','QBR_EVOLPREVCA',
        'QBR_EVOLPREVCAPRCT',
        SaisieCAValPrevSvg,EvolCAValPrevSvg,OkModifSaisieCAValPrevS,OkModifEvolCAPrevS);
 //Qte prev
 Update('EDTSAISIPREVQTE','EDTEVOLVALPREVQTE','QBR_SAISIPREVQTE','QBR_EVOLPREVQTE',
        'QBR_EVOLPREVQTEPRC',
        SaisieQteValPrevSvg,EvolQteValPrevSvg,OkModifSaisieQteValPrevS,OkModifEvolQtePrevS);
 }

 ModifSession(GetControlText('QBR_CODESESSION'),'X');
 ModifSessionVue(GetControlText('QBR_CODESESSION'),'X','X');
end ;


//dans le cas de la duplication de noeud
//il faut qu'on ajoute dans la table
//les différents sous niveaux du noeud à dupliquer
procedure DuplicationSurNiveauSuivant(numnoeudDupli,numNoeudDPere,codeSession,
          valeurAxe,valaxeniv1N,valaxeniv2N,valaxeniv3N,
          valaxeniv4N,valaxeniv5N,valaxeniv6N,
          valaxeniv7N,valaxeniv8N,valaxeniv9N:hString;
          niveau,nivMax:integer;OkRepriseHisto:boolean;
          var numNoeudD:hString);
var Q:TQuery;
    numnoeudX,numnoeudPereX,valeurAxeX:hString;
    histovalX,histoprctX,newvalX,newprctX:double;
    niveauX,codeaxeX:hString;
    valaxeniv1X,valaxeniv2X,valaxeniv3X,valaxeniv4X,valaxeniv5X:hString;
    valaxeniv6X,valaxeniv7X,valaxeniv8X,valaxeniv9X,valBloq:hString;
    QtePrct,QteVal,NewQtePrct,NewQteVal:double;
    CAPrct2,CAVal2,NewCAPrct2,NewCAVal2:double;
    CAPrct3,CAVal3,NewCAPrct3,NewCAVal3:double;
    CAPrct4,CAVal4,NewCAPrct4,NewCAVal4:double;
    CAPrct5,CAVal5,NewCAPrct5,NewCAVal5:double;
    CAPrct6,CAVal6,NewCAPrct6,NewCAVal6:double;
    datedelai:TDateTime;
begin
 if niveau>nivMax
  then exit;

 Q:=MOpenSql('SELECT QBR_NUMNOEUD,'+
             'QBR_NUMNOEUDPERE,QBR_VALEURAXE,QBR_REF1,QBR_REFPRCT1,'+
             'QBR_OP1,QBR_OPPRCT1, '+
             'QBR_NIVEAU,QBR_CODEAXE,QBR_VALAXENIV1,QBR_VALAXENIV2,'+
             'QBR_VALAXENIV3,QBR_VALAXENIV4,QBR_VALAXENIV5,QBR_VALAXENIV6,'+
             'QBR_VALAXENIV7,QBR_VALAXENIV8,QBR_VALAXENIV9,'+
             'QBR_QTEREFPRCT,QBR_QTEREF,QBR_QTECPRCT,QBR_QTEC,'+
             'QBR_REFPRCT2,QBR_REF2,QBR_OPPRCT2,QBR_OP2,'+
             'QBR_REFPRCT3,QBR_REF3,QBR_OPPRCT3,QBR_OP3,'+
             'QBR_REFPRCT4,QBR_REF4,QBR_OPPRCT4,QBR_OP4,'+
             'QBR_REFPRCT5,QBR_REF5,QBR_OPPRCT5,QBR_OP5,'+
             'QBR_REFPRCT6,QBR_REF6,QBR_OPPRCT6,QBR_OP6,QBR_DATEDELAI,'+
             'QBR_VALBLOQUE '+
             ' FROM QBPARBRE '+
             'WHERE QBR_CODESESSION="'+codeSession+
             '" AND QBR_NUMNOEUDPERE="'+numnoeudDupli+'"','',true);
 while not Q.eof do
  begin
   numnoeudX:=Q.fields[0].asString;
   numnoeudPereX:=Q.fields[1].asString;
   valeurAxeX:=Q.fields[2].asString;
   if OkRepriseHisto
    then
     begin
      //CA1
      histovalX:=Q.fields[3].asFloat;
      histoprctX:=Q.fields[4].asFloat;
      newvalX:=Q.fields[5].asFloat;
      newprctX:=Q.fields[6].asFloat;
      //Qte
      QtePrct:=Q.fields[18].asFloat;
      QteVal:=Q.fields[19].asFloat;
      NewQtePrct:=Q.fields[20].asFloat;
      NewQteVal:=Q.fields[21].asFloat;
      //CA2
      CAPrct2:=Q.fields[22].asFloat;
      CAVal2:=Q.fields[23].asFloat;
      NewCAPrct2:=Q.fields[24].asFloat;
      NewCAVal2:=Q.fields[25].asFloat;
      //CA3
      CAPrct3:=Q.fields[26].asFloat;
      CAVal3:=Q.fields[27].asFloat;
      NewCAPrct3:=Q.fields[28].asFloat;
      NewCAVal3:=Q.fields[29].asFloat;
      //CA4
      CAPrct4:=Q.fields[30].asFloat;
      CAVal4:=Q.fields[31].asFloat;
      NewCAPrct4:=Q.fields[32].asFloat;
      NewCAVal4:=Q.fields[33].asFloat;
      //CA5
      CAPrct5:=Q.fields[34].asFloat;
      CAVal5:=Q.fields[35].asFloat;
      NewCAPrct5:=Q.fields[36].asFloat;
      NewCAVal5:=Q.fields[37].asFloat;
      //CA6
      CAPrct6:=Q.fields[38].asFloat;
      CAVal6:=Q.fields[39].asFloat;
      NewCAPrct6:=Q.fields[40].asFloat;
      NewCAVal6:=Q.fields[41].asFloat;
     end
    else
     begin
      //CA1
      histovalX:=0;
      histoprctX:=0;
      newvalX:=Q.fields[5].asFloat;
      newprctX:=Q.fields[6].asFloat;
      //Qte
      QtePrct:=0;
      QteVal:=0;
      NewQtePrct:=0;
      NewQteVal:=Q.fields[21].asFloat;
      //CA2
      CAPrct2:=0;
      CAVal2:=0;
      NewCAPrct2:=Q.fields[24].asFloat;
      NewCAVal2:=Q.fields[25].asFloat;
      //CA3
      CAPrct3:=0;
      CAVal3:=0;
      NewCAPrct3:=Q.fields[28].asFloat;
      NewCAVal3:=Q.fields[29].asFloat;
      //CA4
      CAPrct4:=0;
      CAVal4:=0;
      NewCAPrct4:=Q.fields[32].asFloat;
      NewCAVal4:=Q.fields[33].asFloat;
      //CA5
      CAPrct5:=0;
      CAVal5:=0;
      NewCAPrct5:=Q.fields[36].asFloat;
      NewCAVal5:=Q.fields[37].asFloat;
      //CA6
      CAPrct6:=0;
      CAVal6:=0;
      NewCAPrct6:=Q.fields[40].asFloat;
      NewCAVal6:=Q.fields[41].asFloat;
     end;
   niveauX:=Q.fields[7].asString;
   codeaxeX:=Q.fields[8].asString;
   valaxeniv1X:=Q.fields[9].asString;
   valaxeniv2X:=Q.fields[10].asString;
   valaxeniv3X:=Q.fields[11].asString;
   valaxeniv4X:=Q.fields[12].asString;
   valaxeniv5X:=Q.fields[13].asString;
   valaxeniv6X:=Q.fields[14].asString;
   valaxeniv7X:=Q.fields[15].asString;
   valaxeniv8X:=Q.fields[16].asString;
   valaxeniv9X:=Q.fields[17].asString;

   dateDelai:=Q.fields[42].asDateTime;
   valBloq:=Q.fields[43].asString;

   if valaxeniv1N<>''
    then valaxeniv1X:=valaxeniv1N;
   if valaxeniv2N<>''
    then valaxeniv2X:=valaxeniv2N;
   if valaxeniv3N<>''
    then valaxeniv3X:=valaxeniv3N;
   if valaxeniv4N<>''
    then valaxeniv4X:=valaxeniv4N;
   if valaxeniv5N<>''
    then valaxeniv5X:=valaxeniv5N;
   if valaxeniv6N<>''
    then valaxeniv6X:=valaxeniv6N;
   if valaxeniv7N<>''
    then valaxeniv7X:=valaxeniv7N;
   if valaxeniv8N<>''
    then valaxeniv8X:=valaxeniv8N;
   if valaxeniv9N<>''
    then valaxeniv9X:=valaxeniv9N;

   numNoeudD:=IntToStr(VALEURI(numNoeudD)+1);
   //insertion dans qbpsaisieprev
   MExecuteSql('INSERT INTO QBPARBRE (QBR_CODESESSION,QBR_NUMNOEUD,'+
             'QBR_NUMNOEUDPERE,QBR_VALEURAXE,QBR_REF1,QBR_REFPRCT1,'+
             'QBR_OP1,QBR_OPPRCT1, '+
             'QBR_NIVEAU,QBR_CODEAXE,QBR_VALAXENIV1,QBR_VALAXENIV2,'+
             'QBR_VALAXENIV3,QBR_VALAXENIV4,QBR_VALAXENIV5,QBR_VALAXENIV6,'+
             'QBR_VALAXENIV7,QBR_VALAXENIV8,QBR_VALAXENIV9,'+  
             'QBR_QTEREFPRCT,QBR_QTEREF,QBR_QTECPRCT,QBR_QTEC,'+
             'QBR_REFPRCT2,QBR_REF2,QBR_OPPRCT2,QBR_OP2,'+
             'QBR_REFPRCT3,QBR_REF3,QBR_OPPRCT3,QBR_OP3,'+
             'QBR_REFPRCT4,QBR_REF4,QBR_OPPRCT4,QBR_OP4,'+
             'QBR_REFPRCT5,QBR_REF5,QBR_OPPRCT5,QBR_OP5,'+
             'QBR_REFPRCT6,QBR_REF6,QBR_OPPRCT6,QBR_OP6,QBR_DATEDELAI,'+
             'QBR_VALBLOQUE,QBR_VALBLOQUETMP ) '+
             ' VALUES("'+codeSession+'","'+numNoeudD+'",'+
             '"'+numNoeudDPere+'","'+valeurAxeX+'","'+strFPoint4(histovalX)+
             '","'+strFPoint4(histoprctX)+'",'+
             '"'+strFPoint4(newvalX)+'","'+strFPoint4(newprctX)+'",'+
             '"'+niveauX+'","'+codeaxeX+'","'+valaxeniv1X+'","'+valaxeniv2X+'",'+
             '"'+valaxeniv3X+'","'+valaxeniv4X+'","'+valaxeniv5X+
             '","'+valaxeniv6X+'","'+valaxeniv7X+'","'+valaxeniv8X+
             '","'+valaxeniv9X+
             '","'+strFPoint4(QtePrct)+'","'+strFPoint4(QteVal)+
             '","'+strFPoint4(NewQtePrct)+'","'+strFPoint4(NewQteVal)+
             '","'+strFPoint4(CAPrct2)+'","'+strFPoint4(CAVal2)+
             '","'+strFPoint4(NewCAPrct2)+'","'+strFPoint4(NewCAVal2)+
             '","'+strFPoint4(CAPrct3)+'","'+strFPoint4(CAVal3)+
             '","'+strFPoint4(NewCAPrct3)+'","'+strFPoint4(NewCAVal3)+
             '","'+strFPoint4(CAPrct4)+'","'+strFPoint4(CAVal4)+
             '","'+strFPoint4(NewCAPrct4)+'","'+strFPoint4(NewCAVal4)+
             '","'+strFPoint4(CAPrct5)+'","'+strFPoint4(CAVal5)+
             '","'+strFPoint4(NewCAPrct5)+'","'+strFPoint4(NewCAVal5)+
             '","'+strFPoint4(CAPrct6)+'","'+strFPoint4(CAVal6)+
             '","'+strFPoint4(NewCAPrct6)+'","'+strFPoint4(NewCAVal6)+
             '","'+USDATETIME(dateDelai)+'","'+valBloq+'","'+valBloq+'")',
             'QBPARBRE_TOM (DuplicationSurNiveauSuivant).');

   DuplicationSurNiveauSuivant(numnoeudX,numnoeudD,codeSession,ValeurAxe,
                               valaxeniv1N,valaxeniv2N,valaxeniv3N,
                               valaxeniv4N,valaxeniv5N,valaxeniv6N,
                               valaxeniv7N,valaxeniv8N,valaxeniv9N,
                               niveau+1,nivMax,OkRepriseHisto,numnoeudD);
   Q.next;
  end;
 ferme(Q);
end;

procedure TOM_QBPARBRE.AfterUpdate(okQte:boolean;const ChpCorrectifPrct,EdtSaisie,EdtCorrectif,
          CorrectifPrctSvg,SaisieSvg,CorrectifSvg,ChpNewVal,ChpVal:hString;
          OkModifSaisi,OkModifCorrectif:boolean;
          var codeSql,codeSqlNewValeurNoeud:hString);
var OkModifSaisie,OkModifCorrectifPrct,OkModifCorrectifVal:boolean;
    CorrectifPrct,SaisieVal,CorrectifVal:hString;
begin
  OkModifSaisie:=false;
  OkModifCorrectifPrct:=false;
  OkModifCorrectifVal:=false;
  CorrectifPrct:=FloatToStr(THNumEdit(GetControl(ChpCorrectifPrct)).value);
  SaisieVal:=FloatToStr(VALEUR(GetControlText(EdtSaisie)));
  CorrectifVal:=FloatToStr(VALEUR(GetControlText(EdtCorrectif)));
  if (CorrectifPrct<>CorrectifPrctSvg) then OkModifCorrectifPrct:=true;
  if ((SaisieVal<>SaisieSvg)) and (OkModifSaisi) then
  begin
    OkModifSaisie:=true;
    if (GetControlText(EdtSaisie)='') and (CorrectifPrct<>'0') then
    begin
      OkModifCorrectifPrct:=true;
      OkModifSaisie:=false;
    end;
    if (GetControlText(EdtSaisie)='') and (CorrectifVal<>'') then
    begin
      OkModifCorrectifVal:=true;
      OkModifSaisie:=false;
    end;
  end;
  if (GetControlText(EdtSaisie)<>'') and (OkModifSaisi) then OkModifSaisie:=true;
  if ((CorrectifVal<>CorrectifSvg)) and (OkModifCorrectif) then OkModifCorrectifVal:=true;
  if CorrectifVal='' then CorrectifVal:='0';

  //règle de modification
  //si une valeur saisie est renseignée => on met à jour avec cette valeur
  //si pas renseignée et si evolution renseignée => on met à jour avec aX+b
  if OkModifSaisie then
  begin
    if SaisieVal='' then
    begin
      if codesql='' then
      begin
        codesql:=' '+ChpNewVal+'='+ChpVal+' ';
        codeSqlNewValeurNoeud:= ' '+ChpVal+' ';
      end
      else
      begin
        codeSql:=codesql+' ,'+ChpNewVal+'='+ChpVal+' ';
        if ContextBP=3 then codeSqlNewValeurNoeud:= ' '+ChpVal+' ';
      end;
    end
    else
    begin
      if codesql='' then
      begin
        codeSql:=' '+ChpNewVal+'='+STRFPOINT(VALEUR(SaisieVal))+' ';
        codeSqlNewValeurNoeud:= ' '+STRFPOINT(VALEUR(SaisieVal))+' ';
      end
      else
      begin
        codeSql:=codesql+' ,'+ChpNewVal+'='+STRFPOINT(VALEUR(SaisieVal))+' ';
        if ContextBP=3 then codeSqlNewValeurNoeud:= ' '+STRFPOINT(VALEUR(SaisieVal))+' ';
      end;
    end
  end
  else
  begin
    if OkModifCorrectifPrct then
    begin
      if codesql='' then
      begin
        if okQte then
        begin
          codeSql:=' '+ChpNewVal+'=ROUND('+ChpVal+'+(('+ChpVal+'*'+STRFPOINT(VALEUR(CorrectifPrct))+')/100) + '+STRFPOINT(VALEUR(CorrectifVal))+',0) ';
          codeSqlNewValeurNoeud:=' ROUND('+ChpVal+'+(('+ChpVal+'*'+STRFPOINT(VALEUR(CorrectifPrct))+')/100) + '+STRFPOINT(VALEUR(CorrectifVal))+',0) ';
          { EVI / Flag pour bug pourcentage d'évolution = 0 }
          if VALEUR(CorrectifPrct)=0 then codeSql:=' QBR_VALMODIF="X",'+ codeSql;
        end
        else
        begin
          codeSql:=' '+ChpNewVal+'='+ChpVal+'+(('+ChpVal+'*'+STRFPOINT(VALEUR(CorrectifPrct))+')/100) + '+STRFPOINT(VALEUR(CorrectifVal))+' ';
          codeSqlNewValeurNoeud:=' '+ChpVal+'+(('+ChpVal+'*'+STRFPOINT(VALEUR(CorrectifPrct))+')/100) + '+STRFPOINT(VALEUR(CorrectifVal))+' ';
          { EVI / Flag pour bug pourcentage d'évolution = 0 }
          if VALEUR(CorrectifPrct)=0 then codeSql:=' QBR_VALMODIF="X",'+ codeSql;
        end;
      end
      else
      begin
        if okQte then
        begin
          codeSql:=codesql+' ,'+ChpNewVal+'=ROUND('+ChpVal+'+(('+ChpVal+'*'+STRFPOINT(VALEUR(CorrectifPrct))+')/100) + '+STRFPOINT(VALEUR(CorrectifVal))+',0) ';
          if ContextBP=3 then codeSqlNewValeurNoeud:=' ROUND('+ChpVal+'+(('+ChpVal+'*'+STRFPOINT(VALEUR(CorrectifPrct))+')/100) + '+STRFPOINT(VALEUR(CorrectifVal))+',0) ';
        end
        else
        begin
          codeSql:=codesql+' ,'+ChpNewVal+'='+ChpVal+'+(('+ChpVal+'*'+STRFPOINT(VALEUR(CorrectifPrct))+')/100) + '+STRFPOINT(VALEUR(CorrectifVal))+' ';;
          if ContextBP=3 then codeSqlNewValeurNoeud:=' '+ChpVal+'+(('+ChpVal+'*'+STRFPOINT(VALEUR(CorrectifPrct))+')/100) + '+STRFPOINT(VALEUR(CorrectifVal))+' ';
        end;
      end;
    end;
    if OkModifCorrectifVal and (not OkModifCorrectifPrct) then
    begin
      if codesql='' then
      begin
        if okQte then
        begin
          codeSql:=' '+ChpNewVal+'=ROUND('+ChpVal+'+(('+ChpVal+'*'+STRFPOINT(VALEUR(CorrectifPrct))+')/100) + '+STRFPOINT(VALEUR(CorrectifVal))+',0) ';
          codeSqlNewValeurNoeud:=' ROUND('+ChpVal+'+(('+ChpVal+'*'+STRFPOINT(VALEUR(CorrectifPrct))+')/100) + '+STRFPOINT(VALEUR(CorrectifVal))+',0) ';
        end
        else
        begin
          codeSql:=' '+ChpNewVal+'='+ChpVal+'+(('+ChpVal+'*'+STRFPOINT(VALEUR(CorrectifPrct))+')/100) + '+STRFPOINT(VALEUR(CorrectifVal))+' ';
          codeSqlNewValeurNoeud:=' '+ChpVal+'+(('+ChpVal+'*'+STRFPOINT(VALEUR(CorrectifPrct))+')/100) + '+STRFPOINT(VALEUR(CorrectifVal))+' ';
        end;
      end
      else
      begin
        if okQte then
        begin
          codeSql:=codesql+' ,'+ChpNewVal+'=ROUND('+ChpVal+'+(('+ChpVal+'*'+STRFPOINT(VALEUR(CorrectifPrct))+')/100) + '+STRFPOINT(VALEUR(CorrectifVal))+',0) ';
          if ContextBP=3 then codeSqlNewValeurNoeud:=' ROUND('+ChpVal+'+(('+ChpVal+'*'+STRFPOINT(VALEUR(CorrectifPrct))+')/100) + '+STRFPOINT(VALEUR(CorrectifVal))+',0) ';
        end
        else
        begin
          codeSql:=codesql+' ,'+ChpNewVal+'='+ChpVal+'+(('+ChpVal+'*'+STRFPOINT(VALEUR(CorrectifPrct))+')/100) + '+STRFPOINT(VALEUR(CorrectifVal))+' ';
          if ContextBP=3 then codeSqlNewValeurNoeud:=' '+ChpVal+'+(('+ChpVal+'*'+STRFPOINT(VALEUR(CorrectifPrct))+')/100) + '+STRFPOINT(VALEUR(CorrectifVal))+' ';
        end;
      end;
    end;
  end;
end;

procedure TOM_QBPARBRE.CodeSqlMAJPrct(const codeSession,NoeudPere:hString;
          var CodeSql:hString);
var TotalQte,TotalCa1,TotalCA2,TotalCA3,TotalCA4,TotalCA5,TotalCA6,TotalPrevu,TotalRet,TotalCaRetenu:double;
     codeSqlCA1,codeSqlCA2,codeSqlCA3,codeSqlCA4:hString;
     codeSqlCA5,codeSqlCA6,codeSqlQte:hString;
begin
 ArbreTotalNiv(codeSession,NoeudPere,
               TotalQte,TotalCa1,TotalCA2,TotalCA3,TotalCA4,TotalCA5,TotalCA6,
               TotalPrevu,TotalRet,TotalCaRetenu);
 codeSqlCA1:='';
 codeSqlCA2:='';
 codeSqlCA3:='';
 codeSqlCA4:='';
 codeSqlCA5:='';
 codeSqlCA6:='';
 codeSqlQte:='';
 if TotalCa1<>0
  then codeSqlCA1:='QBR_OPPRCT1=((QBR_OP1/'+strFPoint4(TotalCa1)+')*100) ';
 if TotalQte<>0
  then codeSqlQte:='QBR_QTECPRCT=((QBR_QTEC/'+strFPoint4(TotalQte)+')*100) ';
 if TotalCa2<>0
  then codeSqlCA2:='QBR_OPPRCT2=((QBR_OP2/'+strFPoint4(TotalCa2)+')*100) ';
 if TotalCa3<>0
  then codeSqlCA3:='QBR_OPPRCT3=((QBR_OP3/'+strFPoint4(TotalCa3)+')*100) ';
 if TotalCa4<>0
  then codeSqlCA4:='QBR_OPPRCT4=((QBR_OP4/'+strFPoint4(TotalCa4)+')*100) ';
 if TotalCa5<>0
  then codeSqlCA5:='QBR_OPPRCT5=((QBR_OP5/'+strFPoint4(TotalCa5)+')*100) ';
 if TotalCa6<>0
  then codeSqlCA6:='QBR_OPPRCT6=((QBR_OP6/'+strFPoint4(TotalCa6)+')*100) ';

 if codeSqlCA1<>''
  then codeSql:=codeSqlCA1;
 if codeSqlCA2<>''
  then
   begin
    if codeSql=''
     then codeSql:=codeSqlCA2
     else codeSql:=codeSql+','+codeSqlCA2;
   end;     
 if codeSqlCA3<>''
  then
   begin
    if codeSql=''
     then codeSql:=codeSqlCA3
     else codeSql:=codeSql+','+codeSqlCA3;
   end;
 if codeSqlCA4<>''
  then
   begin
    if codeSql=''
     then codeSql:=codeSqlCA4
     else codeSql:=codeSql+','+codeSqlCA4;
   end;
 if codeSqlCA5<>''
  then
   begin
    if codeSql=''
     then codeSql:=codeSqlCA5
     else codeSql:=codeSql+','+codeSqlCA5;
   end;
 if codeSqlCA6<>''
  then
   begin
    if codeSql=''
     then codeSql:=codeSqlCA6
     else codeSql:=codeSql+','+codeSqlCA6;
   end;
end;

procedure TOM_QBPARBRE.OnAfterUpdateRecord ;
var code,codeSql,noeud,noeudPere,niveau,codeSqlNewValeurNoeud,numnoeudDetail:hString;
    codeSession,codeSqlS,codeChpModif:hString;
    TotalAncQte,TotalAncCA1,TotalAncCA2,TotalAncCA3,TotalAncCA4,TotalAncCA5,TotalAncCA6:double;
    TotalNQte,TotalNCA1,TotalNCA2,TotalNCA3,TotalNCA4,TotalNCA5,TotalNCA6,TotalPrevuN,TotalRetN:double;
    TotalAncQteX,TotalAncCA1X,TotalAncCA2X,TotalAncCA3X,TotalAncCA4X,TotalAncCA5X,TotalAncCA6X,TotalPrevuX,TotalRetX:double;
    NivMax,i:integer;
    numnoeudT,valeurAxe,codePrevu,coderealise:hString;
    valaxeniv1N,valaxeniv2N,valaxeniv3N,valaxeniv4N:hString;
    valaxeniv5N,valaxeniv6N,valaxeniv7N,valaxeniv8N,valaxeniv9N:hString;
    somme,TotalCaRetenuX,TotalCaRetenu:double;
    okModif:boolean;
    okArrondi:boolean;
    codeArrondi:hString;
    TabSqlNewValeurNoeud:array [0..6] of hString;
begin
  Inherited ;
  code:='0';
  codeSql:='';
  codeSqlNewValeurNoeud:='';
  codeChpModif:='';
  okArrondi:=false;

  //Coeff
  OkModifCoeffRetenu:=false;
  NewCoeff:=THNumEdit(GetControl('QBR_COEFFRETENU')).value;
  if (NewCoeff<>CoeffRetenuSvg) then OkModifCoeffRetenu:=true;

  //CA1
  AfterUpdate(false,'QBR_EVOLPRCT1','EDTSAISI1','EDTEVOLVAL1',
              evolprctSvg,saivalSvg,evolvalSvg,'QBR_OP1','QBR_REF1',
              OkModifS,OkModifC,codeSql,codeSqlNewValeurNoeud);
  TabSqlNewValeurNoeud[1]:=codeSqlNewValeurNoeud;
  if codeSqlNewValeurNoeud<>'' then codeChpModif:='1';

  //Qte
  if ContextBP=3 then codeSqlNewValeurNoeud:='';
  AfterUpdate(true,'QBR_EVOLQTEPRCT','EDTSAISIQTE','EDTEVOLVALQTE',
              EvolQtePrctSvg,SaisieQteValSvg,EvolQteValSvg,'QBR_QTEC','QBR_QTEREF',
              OkModifSaisieQteValS,OkModifEvolQteS,codeSql,codeSqlNewValeurNoeud);
  TabSqlNewValeurNoeud[0]:=codeSqlNewValeurNoeud;
  if (codeSqlNewValeurNoeud<>'') and (codeChpModif='') then
  begin
    codeChpModif:='qte';
    okArrondi:=true;
  end;

  //CA2
  if ContextBP=3 then codeSqlNewValeurNoeud:='';
  AfterUpdate(false,'QBR_EVOLPRCT2','EDTSAISI2','EDTEVOLVAL2',
              EvolCAPrct2Svg,SaisieCAVal2Svg,EvolCAVal2Svg,'QBR_OP2','QBR_REF2',
              OkModifSaisieCAVal2S,OkModifEvolCA2S,codeSql,codeSqlNewValeurNoeud);
  TabSqlNewValeurNoeud[2]:=codeSqlNewValeurNoeud;
  if (codeSqlNewValeurNoeud<>'') and (codeChpModif='') then codeChpModif:='2';

  //CA3
  if ContextBP=3 then codeSqlNewValeurNoeud:='';
  AfterUpdate(false,'QBR_EVOLPRCT3','EDTSAISI3','EDTEVOLVAL3',
              EvolCAPrct3Svg,SaisieCAVal3Svg,EvolCAVal3Svg,'QBR_OP3','QBR_REF3',
              OkModifSaisieCAVal3S,OkModifEvolCA3S,codeSql,codeSqlNewValeurNoeud);
  TabSqlNewValeurNoeud[3]:=codeSqlNewValeurNoeud;
  if (codeSqlNewValeurNoeud<>'') and (codeChpModif='') then codeChpModif:='3';

  //CA4
  if ContextBP=3 then codeSqlNewValeurNoeud:='';
  AfterUpdate(false,'QBR_EVOLPRCT4','EDTSAISI4','EDTEVOLVAL4',
              EvolCAPrct4Svg,SaisieCAVal4Svg,EvolCAVal4Svg,'QBR_OP4','QBR_REF4',
              OkModifSaisieCAVal4S,OkModifEvolCA4S,codeSql,codeSqlNewValeurNoeud);
  TabSqlNewValeurNoeud[4]:=codeSqlNewValeurNoeud;
  if (codeSqlNewValeurNoeud<>'') and (codeChpModif='') then codeChpModif:='4';

  //CA5
  if ContextBP=3 then codeSqlNewValeurNoeud:='';
  AfterUpdate(false,'QBR_EVOLPRCT5','EDTSAISI5','EDTEVOLVAL5',
              EvolCAPrct5Svg,SaisieCAVal5Svg,EvolCAVal5Svg,'QBR_OP5','QBR_REF5',
              OkModifSaisieCAVal5S,OkModifEvolCA5S,codeSql,codeSqlNewValeurNoeud);
  TabSqlNewValeurNoeud[5]:=codeSqlNewValeurNoeud;
  if (codeSqlNewValeurNoeud<>'') and (codeChpModif='') then codeChpModif:='5';

  //CA6
  if ContextBP=3 then codeSqlNewValeurNoeud:='';
  AfterUpdate(false,'QBR_EVOLPRCT6','EDTSAISI6','EDTEVOLVAL6',
              EvolCAPrct6Svg,SaisieCAVal6Svg,EvolCAVal6Svg,'QBR_OP6','QBR_REF6',
              OkModifSaisieCAVal6S,OkModifEvolCA6S,codeSql,codeSqlNewValeurNoeud);
  TabSqlNewValeurNoeud[6]:=codeSqlNewValeurNoeud;
  if (codeSqlNewValeurNoeud<>'') and (codeChpModif='') then codeChpModif:='6';

  { EVI / ORLI : Onglets CA & QTE Prevision }
  {//CA prev
   AfterUpdate(false,'QBR_EVOLPREVCAPRCT','EDTSAISIPREVCA','EDTEVOLVALPREVCA',
               EvolCAPrctPrev,SaisieCAValPrevSvg,EvolCAValPrevSvg,'QBR_CARETENU','QBR_PREVUCA',
               OkModifSaisieCAValPrevS,OkModifEvolCAPrevS,codeSql,codeSqlNewValeurNoeud);

   //Qte prev
   AfterUpdate(true,'QBR_EVOLPREVQTEPRC','EDTSAISIPREVQTE','EDTEVOLVALPREVQTE',
               EvolQtePrctPrev,SaisieQteValPrevSvg,EvolQteValPrevSvg,'QBR_QTERETENUE','QBR_PREVU',
               OkModifSaisieQteValPrevS,OkModifEvolQtePrevS,codeSql,codeSqlNewValeurNoeud);
  }

  if OkInitCoeff then okArrondi:=true;
  codeArrondi:='';
  //recupere les valeurs
  codeSession:=GetField('QBR_CODESESSION');
  noeudPere:=GetField('QBR_NUMNOEUDPERE');
  noeud:=GetField('QBR_NUMNOEUD');
  niveau:=GetField('QBR_NIVEAU');

  //coeff
  if OkModifCoeffRetenu then
  begin
    if codeSql<>'' then codeSql:=codeSql+',';
    if DonneParamS(ps_BPCoeffPerCAQte)='2' then
    begin
      codePrevu:='QBR_PREVU';
      coderealise:='QBR_REALISE';
    end
    else
    begin
      codePrevu:='QBR_PREVUCA';
      coderealise:='QBR_REALISECA';
    end;

    if okArrondi then codeArrondi:=' QBR_PREVU=ROUND((QBR_REALISE*'+STRFPOINT(NewCoeff)+
                                   '),0),QBR_QTERETENUE=ROUND((QBR_REALISE*'+STRFPOINT(NewCoeff)+
                                   '),0),QBR_QTEC=ROUND((QBR_QTEREF*'+STRFPOINT(NewCoeff)+
                                   '),0),QBR_PREVUCA=(QBR_REALISECA*'+STRFPOINT(NewCoeff)+
                                   '),QBR_CARETENU=(QBR_REALISECA*'+STRFPOINT(NewCoeff)
                 else codeArrondi:=' QBR_PREVU=(QBR_REALISE*'+STRFPOINT(NewCoeff)+
                                   '),QBR_QTERETENUE=(QBR_REALISE*'+STRFPOINT(NewCoeff)+
                                   '),QBR_QTEC=(QBR_QTEREF*'+STRFPOINT(NewCoeff)+
                                   '),QBR_PREVUCA=(QBR_REALISECA*'+STRFPOINT(NewCoeff)+
                                   '),QBR_CARETENU=(QBR_REALISECA*'+STRFPOINT(NewCoeff);

    codeSql:=codesql+codeArrondi+'),QBR_OP1=QBR_REF1*'+STRFPOINT(NewCoeff)+
                                 ',QBR_OP2=QBR_REF2*'+STRFPOINT(NewCoeff)+
                                 ',QBR_OP3=QBR_REF3*'+STRFPOINT(NewCoeff)+
                                 ',QBR_OP4=QBR_REF4*'+STRFPOINT(NewCoeff)+
                                 ',QBR_OP5=QBR_REF5*'+STRFPOINT(NewCoeff)+
                                 ',QBR_OP6=QBR_REF6*'+STRFPOINT(NewCoeff)+' ';
  end;

  NivMax:=ChercheNivMaxSession(codeSession);

  if (codeSql='') and (OkInitPrev) then
  begin
    somme:=THNumEdit(GetControl('QBR_QTET1')).value+
           THNumEdit(GetControl('QBR_QTET2')).value+
           THNumEdit(GetControl('QBR_QTET3')).value+
           THNumEdit(GetControl('QBR_QTET4')).value+
           THNumEdit(GetControl('QBR_QTET5')).value+
           THNumEdit(GetControl('QBR_QTET6')).value+
           THNumEdit(GetControl('QBR_QTET7')).value+
           THNumEdit(GetControl('QBR_QTET8')).value+
           THNumEdit(GetControl('QBR_QTET9')).value+
           THNumEdit(GetControl('QBR_QTET10')).value+
           THNumEdit(GetControl('QBR_QTET11')).value+
           THNumEdit(GetControl('QBR_QTET12')).value+
           THNumEdit(GetControl('QBR_QTET13')).value+
           THNumEdit(GetControl('QBR_QTET14')).value+
           THNumEdit(GetControl('QBR_QTET15')).value+
           THNumEdit(GetControl('QBR_QTET16')).value+
           THNumEdit(GetControl('QBR_QTET17')).value+
           THNumEdit(GetControl('QBR_QTET18')).value+
           THNumEdit(GetControl('QBR_QTET19')).value+
           THNumEdit(GetControl('QBR_QTET20')).value;
    if somme<>THNumEdit(GetControl('QBR_QTERETENUE')).value then codeSql:='QBR_QTERETENUE="'+STRFPOINT(somme)+'" ';
  end;


  if (codeSql<>'') then
  begin
    if ContextBP in [0,1,2] then OkModif:=OkModifNoeudArbre(codeSession,noeud,noeudPere,codeSqlNewValeurNoeud,
                                                           codeChpModif,VALEURI(niveau),OkModifValBloque)
    else
    begin
      OkModif:=true;
      For i:= 0 to (NbrValAff(GetField('QBR_CODESESSION')))-1 do
      begin
        if TabSqlNewValeurNoeud[i]<>'' then
        begin
          if i=0 then codeChpModif:='qte' else codeChpModif:=IntToStr(i);
          OkModif:=OkModifNoeudArbre(codeSession,noeud,noeudPere,TabSqlNewValeurNoeud[i],
                                     codeChpModif,VALEURI(niveau),OkModifValBloque);
          if OkModif=false then break;
        end;
      end;
    end;

    OkModifValBloque:=false;

    if OkModif then
    begin
      //total valeur niveau
      //on recupère les valeurs Histo avant modification
      ArbreTotalNivRef(codeSession,noeudPere,
                       TotalAncQte,TotalAncCA1,TotalAncCA2,TotalAncCA3,
                       TotalAncCA4,TotalAncCA5,TotalAncCA6);
      //on recupère les valeurs courante avant modification
      ArbreTotalNiv(codeSession,noeudPere,
                    TotalAncQteX,TotalAncCA1X,TotalAncCA2X,TotalAncCA3X,
                    TotalAncCA4X,TotalAncCA5X,TotalAncCA6X,TotalPrevuX,TotalRetX,TotalCaRetenuX);

      //maj du noeud
      MExecuteSql('UPDATE QBPARBRE SET '+codeSql+
                  ' WHERE QBR_CODESESSION="'+codeSession+
                  '" AND QBR_NUMNOEUD="'+Noeud+'"',
                  'QBPARBRE_TOM (OnAfterUpdateRecord).');

      ArbreTotalNiv(codeSession,noeudPere,
                    TotalNQte,TotalNCA1,TotalNCA2,TotalNCA3,TotalNCA4,
                    TotalNCA5,TotalNCA6,TotalPrevuN,TotalRetN,TotalCaRetenu);

      MiseAjourNiv(true,codeSession,Noeud,noeudPere,niveau,nivMaxSession,
                   TotalNQte,TotalNCA1,TotalNCA2,TotalNCA3,
                   TotalNCA4,TotalNCA5,TotalNCA6,
                   TotalNQte-TotalAncQte,TotalNCA1-TotalAncCA1,
                   TotalNCA2-TotalAncCA2,TotalNCA3-TotalAncCA3,
                   TotalNCA4-TotalAncCA4,TotalNCA5-TotalAncCA5,
                   TotalNCA6-TotalAncCA6,
                   STRFPOINT(NewCoeff),OkModifCoeffRetenu,OkInitPrev,OkInitNivTaille);

      { EVI / Gestion des axes ajoutés manuellement
        - si ce n'est pas une session en Loi d'Eclatement
        - si l'axe ajouté/modifié correspond au dernier niveau de l'arbre (càd à une date)
        - si il ne s'agit pas d'une modification d'un axe généré automatiquement
        alors on effectue une copie de la ligne de QBPARBRE dans QBPARBREDETAIL ou on la modifie
      }

      if DonneMethodeSession(codeSession) = '1' then
      begin
        if StrToInt(niveau)=NivMax+1 then
        begin
          if noeud > noeudPere then
          begin
            { GC/GRC : EV / Correction valeurs doublées, vérification qu'il ne s'agit pas d'une modification d'un noeud existant }
            if not ExisteSQL('SELECT QBH_NUMNOEUDREF FROM QBPARBREDETAIL WHERE QBH_CODESESSION="'+codeSession+
                             '" AND QBH_NUMNOEUDREF="'+Noeud+'"')
            then
            begin
              numnoeudDetail:=IntToStr(BPIncrementeNumNoeudDetail(codeSession));
              MExecuteSql('INSERT INTO QBPARBREDETAIL (QBH_CODESESSION,QBH_NUMNOEUD,QBH_NIVEAU,'+
                          'QBH_DATEDELAI,'+
                          'QBH_OP1,QBH_OPPRCT1,QBH_REF1,QBH_REFPRCT1,'+
                          'QBH_OP2,QBH_OPPRCT2,QBH_REF2,QBH_REFPRCT2,'+
                          'QBH_OP3,QBH_OPPRCT3,QBH_REF3,QBH_REFPRCT3,'+
                          'QBH_OP4,QBH_OPPRCT4,QBH_REF4,QBH_REFPRCT4,'+
                          'QBH_OP5,QBH_OPPRCT5,QBH_REF5,QBH_REFPRCT5,'+
                          'QBH_OP6,QBH_OPPRCT6,QBH_REF6,QBH_REFPRCT6,'+
                          'QBH_QTEC,QBH_QTECPRCT,QBH_QTEREF,QBH_QTEREFPRCT,'+
                          'QBH_DATEPIECE,QBH_NUMNOEUDREF) SELECT QBR_CODESESSION,'+numnoeudDetail+',QBR_NIVEAU,'+
                          'QBR_DATEDELAI,'+
                          'QBR_OP1,QBR_OPPRCT1,QBR_REF1,QBR_REFPRCT1,'+
                          'QBR_OP2,QBR_OPPRCT2,QBR_REF2,QBR_REFPRCT2,'+
                          'QBR_OP3,QBR_OPPRCT3,QBR_REF3,QBR_REFPRCT3,'+
                          'QBR_OP4,QBR_OPPRCT4,QBR_REF4,QBR_REFPRCT4,'+
                          'QBR_OP5,QBR_OPPRCT5,QBR_REF5,QBR_REFPRCT5,'+
                          'QBR_OP6,QBR_OPPRCT6,QBR_REF6,QBR_REFPRCT6,'+
                          'QBR_QTEC,QBR_QTECPRCT,QBR_QTEREF,QBR_QTEREFPRCT,'+
                          'QBR_DATEDELAI,QBR_NUMNOEUD'+
                          ' FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+'" AND QBR_NUMNOEUD="'+Noeud+'"',
                          'QBPARBRE_TOM (OnAfterUpdateRecord).');
            end
          end;
        end;
      end;
    end
    else
    begin
      { EVI / Contrôle de QBR_VALBLOQUE }
      if TCheckBox(GetControl('QBR_VALBLOQUE')).checked =false
      then PGIINFO('Valeur non modifiable (blocage des valeurs des autres niveaux).');
    end;
  end;

  if OkNew then
  begin
    CodeSqlMAJPrct(codeSession,NoeudPere,codeSqlS);
    if codeSqlS<>'' then MExecuteSql('UPDATE QBPARBRE SET '+codeSqlS+
                                     ' WHERE QBR_CODESESSION="'+codeSession+
                                     '" AND QBR_NUMNOEUDPERE="'+NoeudPere+'"',
                                     'QBPARBRE_TOM (OnAfterUpdateRecord).');
  end;

  if OkDuplication then
  begin
    numnoeudT:=IntToStr(BPIncrementenumNoeud(codeSession));
    valeurAxe:=GetField('QBR_VALEURAXE');
    case ValeurI(niveau) of
     1 : valaxeniv1N:=valeurAxe;
     2 : valaxeniv2N:=valeurAxe;
     3 : valaxeniv3N:=valeurAxe;
     4 : valaxeniv4N:=valeurAxe;
     5 : valaxeniv5N:=valeurAxe;
     6 : valaxeniv6N:=valeurAxe;
     7 : valaxeniv7N:=valeurAxe;
     8 : valaxeniv8N:=valeurAxe;
     9 : valaxeniv9N:=valeurAxe;
    end;

    DuplicationSurNiveauSuivant(numNoeudDup,noeud,codeSession,
                                valeurAxe,valaxeniv1N,valaxeniv2N,valaxeniv3N,
                                valaxeniv4N,valaxeniv5N,valaxeniv6N,
                                valaxeniv7N,valaxeniv8N,valaxeniv9N,
                                VALEURI(niveau),NivMax,
                                OkrepriseHisto,noeud);
    OkDuplication:=false;

    ArbreTotalNiv(codeSession,noeudPere,
                  TotalNQte,TotalNCA1,TotalNCA2,TotalNCA3,TotalNCA4,
                  TotalNCA5,TotalNCA6,TotalPrevuN,TotalRetN,TotalCaRetenu);

    MiseAjourNiv(true,codeSession,Noeud,noeudPere,niveau,nivMaxSession,
                 TotalNQte,TotalNCA1,TotalNCA2,TotalNCA3,
                 TotalNCA4,TotalNCA5,TotalNCA6,
                 TotalNQte-TotalAncQte,TotalNCA1-TotalAncCA1,
                 TotalNCA2-TotalAncCA2,TotalNCA3-TotalAncCA3,
                 TotalNCA4-TotalAncCA4,TotalNCA5-TotalAncCA5,
                 TotalNCA6-TotalAncCA6,
                 STRFPOINT(NewCoeff),OkModifCoeffRetenu,OkInitPrev,OkInitNivTaille);

  end;

  MAJPrctVariation(codeSession);
  THEdit(GetControl('QBR_VALEURAXE')).Enabled:=false;
  THCheckBox(GetControl('CB_NEWVALUE')).Visible:=false;
end ;


procedure TOM_QBPARBRE.AffichageColonneListe;
var TabActive,CAQte:hString;
    i,NivMax:integer;
    LgColValeur,LgColRef,LgColOP,LgColRefPrct,LgColOPPrct:integer;
    LgPrctVariation,LgColCache:integer;
begin
 TabActive:=THPageControl2(GetControl('PAGE' )).ActivePage.Name;
 LgColValeur:=120;
 LgColRef:=120;
 LgColOP:=120;
 LgColRefPrct:=120;
 LgColOPPrct:=120;
 LgPrctVariation:=101;
 LgColCache:=-1;

 NivMax:=ChercheNivMax(getfield('QBR_CODESESSION'));
 
 CAQte:='CA ';
 if DonneParamS(ps_BPCoeffPerCAQte)='2'
  then CAQte:='Qté ';

 // if (SessionInitPrev(getfield('QBR_CODESESSION'))) and (VALEUR(niveauT)>ChercheNivMax(getfield('QBR_CODESESSION')))

 for i:=1 to TF.LaGrid.colcount do
  begin
  if OkInitObj
  then
   begin
    if TF.LaGrid.ColNames[i]='QBR_OP1'
     then TF.LaGrid.Cells[i,0]:='Objectif';
    if TF.LaGrid.ColNames[i]='QBR_OP2'
     then TF.LaGrid.Cells[i,0]:='Objectif';
    if TF.LaGrid.ColNames[i]='QBR_OP3'
     then TF.LaGrid.Cells[i,0]:='Objectif';
    if TF.LaGrid.ColNames[i]='QBR_OP4'
     then TF.LaGrid.Cells[i,0]:='Objectif';
    if TF.LaGrid.ColNames[i]='QBR_OP5'
     then TF.LaGrid.Cells[i,0]:='Objectif';
    if TF.LaGrid.ColNames[i]='QBR_OP6'
     then TF.LaGrid.Cells[i,0]:='Objectif';
    if TF.LaGrid.ColNames[i]='QBR_QTEC'
     then TF.LaGrid.Cells[i,0]:='Objectif';
   end
  else
   begin
    if TF.LaGrid.ColNames[i]='QBR_OP1'
     then TF.LaGrid.Cells[i,0]:='Prévu';
    if TF.LaGrid.ColNames[i]='QBR_OP2'
     then TF.LaGrid.Cells[i,0]:='Prévu';
    if TF.LaGrid.ColNames[i]='QBR_OP3'
     then TF.LaGrid.Cells[i,0]:='Prévu';
    if TF.LaGrid.ColNames[i]='QBR_OP4'
     then TF.LaGrid.Cells[i,0]:='Prévu';
    if TF.LaGrid.ColNames[i]='QBR_OP5'
     then TF.LaGrid.Cells[i,0]:='Prévu';
    if TF.LaGrid.ColNames[i]='QBR_OP6'
     then TF.LaGrid.Cells[i,0]:='Prévu';
    if TF.LaGrid.ColNames[i]='QBR_QTEC'
     then TF.LaGrid.Cells[i,0]:='Prévu';
   end;
  end;

//CA1
 if TabActive='TAB1'
  then
   for i:=1 to TF.LaGrid.colcount do
     begin
      if TF.LaGrid.ColNames[i]='QBR_VALEURAXE'
       then TF.LaGrid.ColWidths[i]:=LgColValeur;
      
      //CA1
      if TF.LaGrid.ColNames[i]='QBR_REF1'
       then TF.LaGrid.ColWidths[i]:=LgColRef;
      if TF.LaGrid.ColNames[i]='QBR_OP1'
       then TF.LaGrid.ColWidths[i]:=LgColOP;
      if TF.LaGrid.ColNames[i]='QBR_REFPRCT1'
       then TF.LaGrid.ColWidths[i]:=LgColRefPrct;
      if TF.LaGrid.ColNames[i]='QBR_OPPRCT1'
       then TF.LaGrid.ColWidths[i]:=LgColOPPrct;
      //CA2
      if TF.LaGrid.ColNames[i]='QBR_REF2'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OP2'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_REFPRCT2'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OPPRCT2'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      //CA3
      if TF.LaGrid.ColNames[i]='QBR_REF3'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OP3'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_REFPRCT3'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OPPRCT3'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      //QTE
      if TF.LaGrid.ColNames[i]='QBR_QTEREF'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_QTEREFPRCT'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_QTEC'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_QTECPRCT'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      //Coeff
      if TF.LaGrid.ColNames[i]='QBR_HISTO'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_REALISE'
          then TF.LaGrid.ColWidths[i]:=LgColCache;
         if TF.LaGrid.ColNames[i]='QBR_PREVU'
          then TF.LaGrid.ColWidths[i]:=LgColCache;
         if TF.LaGrid.ColNames[i]='QBR_COEFFCALCUL'
          then TF.LaGrid.ColWidths[i]:=LgColCache;
         if TF.LaGrid.ColNames[i]='QBR_COEFFRETENU'
          then TF.LaGrid.ColWidths[i]:=LgColCache;      
      if TF.LaGrid.ColNames[i]='QBR_QTERETENUE'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_CARETENU'
       then TF.LaGrid.ColWidths[i]:=LgColCache;

         if TF.LaGrid.ColNames[i]='QBR_REALISE'
          then TF.LaGrid.Cells[i,0]:='Réalisé';
         if TF.LaGrid.ColNames[i]='QBR_PREVU'
          then TF.LaGrid.Cells[i,0]:='Prévu';

         if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATION1'
          then TF.LaGrid.ColWidths[i]:=LgColCache;
         if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATION2'
          then TF.LaGrid.ColWidths[i]:=LgColCache;
         if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATION3'
          then TF.LaGrid.ColWidths[i]:=LgColCache;
         if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATIONQ'
          then TF.LaGrid.ColWidths[i]:=LgPrctVariation;
    end;

 //Qte
 if TabActive='TAB2'
  then
   begin
    for i:=1 to TF.LaGrid.colcount do
     begin
      if TF.LaGrid.ColNames[i]='QBR_VALEURAXE'
       then TF.LaGrid.ColWidths[i]:=LgColValeur;
      //CA1
      if TF.LaGrid.ColNames[i]='QBR_REF1'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OP1'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_REFPRCT1'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OPPRCT1'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      //CA2
      if TF.LaGrid.ColNames[i]='QBR_REF2'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OP2'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_REFPRCT2'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OPPRCT2'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      //CA3
      if TF.LaGrid.ColNames[i]='QBR_REF3'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OP3'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_REFPRCT3'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OPPRCT3'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      //QTE
      if TF.LaGrid.ColNames[i]='QBR_QTEREF'
       then TF.LaGrid.ColWidths[i]:=LgColRef;
      if TF.LaGrid.ColNames[i]='QBR_QTEREFPRCT'
       then TF.LaGrid.ColWidths[i]:=LgColRef;
      if TF.LaGrid.ColNames[i]='QBR_QTEC'
       then TF.LaGrid.ColWidths[i]:=LgColRef;
      if TF.LaGrid.ColNames[i]='QBR_QTECPRCT'
       then TF.LaGrid.ColWidths[i]:=LgColRef;
      //Coeff
      if TF.LaGrid.ColNames[i]='QBR_HISTO'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_REALISE'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_PREVU'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_COEFFCALCUL'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_COEFFRETENU'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_QTERETENUE'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_HISTOCA'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_PREVUCA'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_REALISECA'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_CARETENU'
       then TF.LaGrid.ColWidths[i]:=LgColCache;

      if TF.LaGrid.ColNames[i]='QBR_REALISE'
       then TF.LaGrid.Cells[i,0]:='Réalisé';
      if TF.LaGrid.ColNames[i]='QBR_PREVU'
       then TF.LaGrid.Cells[i,0]:='Prévu';

      if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATION1'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATION2'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATION3'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATIONQ'
       then TF.LaGrid.ColWidths[i]:=LgPrctVariation;
     end;
   end;

 //CA2
 if TabActive='TAB3'
  then
   for i:=1 to TF.LaGrid.colcount do
     begin
      if TF.LaGrid.ColNames[i]='QBR_VALEURAXE'
       then TF.LaGrid.ColWidths[i]:=LgColValeur;
      //CA1
      if TF.LaGrid.ColNames[i]='QBR_REF1'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OP1'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_REFPRCT1'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OPPRCT1'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      //CA2
      if TF.LaGrid.ColNames[i]='QBR_REF2'
       then TF.LaGrid.ColWidths[i]:=LgColRef;
      if TF.LaGrid.ColNames[i]='QBR_OP2'
       then TF.LaGrid.ColWidths[i]:=LgColOP;
      if TF.LaGrid.ColNames[i]='QBR_REFPRCT2'
       then TF.LaGrid.ColWidths[i]:=LgColRefPrct;
      if TF.LaGrid.ColNames[i]='QBR_OPPRCT2'
       then TF.LaGrid.ColWidths[i]:=LgColOPPrct;
      //CA3
      if TF.LaGrid.ColNames[i]='QBR_REF3'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OP3'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_REFPRCT3'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OPPRCT3'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      //QTE
      if TF.LaGrid.ColNames[i]='QBR_QTEREF'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_QTEREFPRCT'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_QTEC'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_QTECPRCT'
       then TF.LaGrid.ColWidths[i]:=LgColCache;    
      //Coeff
      if TF.LaGrid.ColNames[i]='QBR_HISTO'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_REALISE'
       then TF.LaGrid.ColWidths[i]:=LgColCache;   
      if TF.LaGrid.ColNames[i]='QBR_PREVU'
       then TF.LaGrid.ColWidths[i]:=LgColCache;   
      if TF.LaGrid.ColNames[i]='QBR_COEFFCALCUL'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_COEFFRETENU'
       then TF.LaGrid.ColWidths[i]:=LgColCache;     
      if TF.LaGrid.ColNames[i]='QBR_QTERETENUE'
       then TF.LaGrid.ColWidths[i]:=LgColCache;  
      if TF.LaGrid.ColNames[i]='QBR_HISTOCA'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_PREVUCA'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_REALISECA'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_CARETENU'
       then TF.LaGrid.ColWidths[i]:=LgColCache;

      if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATION1'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATION2'
       then TF.LaGrid.ColWidths[i]:=LgPrctVariation;
      if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATION3'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATIONQ'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
    end;

 //CA3
 if TabActive='TAB4'
  then
   for i:=1 to TF.LaGrid.colcount do
     begin
      if TF.LaGrid.ColNames[i]='QBR_VALEURAXE'
       then TF.LaGrid.ColWidths[i]:=LgColValeur;
      //CA1
      if TF.LaGrid.ColNames[i]='QBR_REF1'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OP1'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_REFPRCT1'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OPPRCT1'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      //CA2
      if TF.LaGrid.ColNames[i]='QBR_REF2'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OP2'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_REFPRCT2'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OPPRCT2'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      //CA3
      if TF.LaGrid.ColNames[i]='QBR_REF3'
       then TF.LaGrid.ColWidths[i]:=LgColRef;
      if TF.LaGrid.ColNames[i]='QBR_OP3'
       then TF.LaGrid.ColWidths[i]:=LgColOP;
      if TF.LaGrid.ColNames[i]='QBR_REFPRCT3'
       then TF.LaGrid.ColWidths[i]:=LgColRefPrct;
      if TF.LaGrid.ColNames[i]='QBR_OPPRCT3'
       then TF.LaGrid.ColWidths[i]:=LgColOPPrct;
      //QTE
      if TF.LaGrid.ColNames[i]='QBR_QTEREF'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_QTEREFPRCT'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_QTEC'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_QTECPRCT'
       then TF.LaGrid.ColWidths[i]:=LgColCache;      
      //Coeff
      if TF.LaGrid.ColNames[i]='QBR_HISTO'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_REALISE'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_PREVU'
       then TF.LaGrid.ColWidths[i]:=LgColCache;   
      if TF.LaGrid.ColNames[i]='QBR_COEFFCALCUL'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_COEFFRETENU'
       then TF.LaGrid.ColWidths[i]:=LgColCache;     
      if TF.LaGrid.ColNames[i]='QBR_QTERETENUE'
       then TF.LaGrid.ColWidths[i]:=LgColCache;   
      if TF.LaGrid.ColNames[i]='QBR_HISTOCA'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_PREVUCA'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_REALISECA'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_CARETENU'
       then TF.LaGrid.ColWidths[i]:=LgColCache;

      if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATION1'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATION2'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATION3'
       then TF.LaGrid.ColWidths[i]:=LgPrctVariation;
      if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATIONQ'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
    end;

 //Coeff
 if TabActive='TAB8'
  then
   for i:=1 to TF.LaGrid.colcount do
     begin
      if TF.LaGrid.ColNames[i]='QBR_VALEURAXE'
       then TF.LaGrid.ColWidths[i]:=LgColValeur;
      //CA1
      if TF.LaGrid.ColNames[i]='QBR_REF1'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OP1'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_REFPRCT1'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OPPRCT1'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      //CA2
      if TF.LaGrid.ColNames[i]='QBR_REF2'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OP2'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_REFPRCT2'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OPPRCT2'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      //CA3
      if TF.LaGrid.ColNames[i]='QBR_REF3'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OP3'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_REFPRCT3'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OPPRCT3'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      //QTE
      if TF.LaGrid.ColNames[i]='QBR_QTEREF'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_QTEREFPRCT'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_QTEC'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_QTECPRCT'
       then TF.LaGrid.ColWidths[i]:=LgColCache;   
      //Coeff
      if DonneParamS(ps_BPCoeffPerCAQte)='2'
       then
        begin
         if TF.LaGrid.ColNames[i]='QBR_HISTO'
          then TF.LaGrid.ColWidths[i]:=LgColOP;
         if TF.LaGrid.ColNames[i]='QBR_PREVU'
          then TF.LaGrid.ColWidths[i]:=LgColOP;
         if TF.LaGrid.ColNames[i]='QBR_REALISE'
          then TF.LaGrid.ColWidths[i]:=LgColOP;
         if TF.LaGrid.ColNames[i]='QBR_HISTOCA'
          then TF.LaGrid.ColWidths[i]:=LgColCache;
         if TF.LaGrid.ColNames[i]='QBR_PREVUCA'
          then TF.LaGrid.ColWidths[i]:=LgColCache;
         if TF.LaGrid.ColNames[i]='QBR_REALISECA'
          then TF.LaGrid.ColWidths[i]:=LgColCache;
        end
       else
        begin
         if TF.LaGrid.ColNames[i]='QBR_HISTO'
          then TF.LaGrid.ColWidths[i]:=LgColCache;
         if TF.LaGrid.ColNames[i]='QBR_PREVU'
          then TF.LaGrid.ColWidths[i]:=LgColCache;
         if TF.LaGrid.ColNames[i]='QBR_REALISE'
          then TF.LaGrid.ColWidths[i]:=LgColCache;
         if TF.LaGrid.ColNames[i]='QBR_HISTOCA'
          then TF.LaGrid.ColWidths[i]:=LgColOP;
         if TF.LaGrid.ColNames[i]='QBR_PREVUCA'
          then TF.LaGrid.ColWidths[i]:=LgColOP;
         if TF.LaGrid.ColNames[i]='QBR_REALISECA'
          then TF.LaGrid.ColWidths[i]:=LgColOP;
        end;
      if TF.LaGrid.ColNames[i]='QBR_COEFFCALCUL'
       then TF.LaGrid.ColWidths[i]:=LgColOP;
      if TF.LaGrid.ColNames[i]='QBR_COEFFRETENU'
       then TF.LaGrid.ColWidths[i]:=LgColOP;
      if TF.LaGrid.ColNames[i]='QBR_QTERETENUE'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_CARETENU'
       then TF.LaGrid.ColWidths[i]:=LgColCache;

      if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATION1'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATION2'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATION3'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATIONQ'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
    end;
 

 //Quantite prev
 if TabActive='TAB10'
  then
   for i:=1 to TF.LaGrid.colcount do
     begin
      if TF.LaGrid.ColNames[i]='QBR_VALEURAXE'
       then TF.LaGrid.ColWidths[i]:=LgColValeur;
      //CA1
      if TF.LaGrid.ColNames[i]='QBR_REF1'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OP1'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_REFPRCT1'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OPPRCT1'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      //CA2
      if TF.LaGrid.ColNames[i]='QBR_REF2'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OP2'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_REFPRCT2'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OPPRCT2'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      //CA3
      if TF.LaGrid.ColNames[i]='QBR_REF3'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OP3'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_REFPRCT3'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OPPRCT3'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      //QTE
      if TF.LaGrid.ColNames[i]='QBR_QTEREF'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_QTEREFPRCT'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_QTEC'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_QTECPRCT'
       then TF.LaGrid.ColWidths[i]:=LgColCache;      
      //Coeff
      if TF.LaGrid.ColNames[i]='QBR_REALISE'
       then TF.LaGrid.ColWidths[i]:=LgColValeur;
      if TF.LaGrid.ColNames[i]='QBR_PREVU'
       then TF.LaGrid.ColWidths[i]:=LgColValeur;
      if TF.LaGrid.ColNames[i]='QBR_COEFFCALCUL'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_QTERETENUE'
       then TF.LaGrid.ColWidths[i]:=LgColValeur;
      if TF.LaGrid.ColNames[i]='QBR_CARETENU'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_HISTOCA'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_PREVUCA'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_REALISECA'
       then TF.LaGrid.ColWidths[i]:=LgColCache;

      if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATION1'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATION2'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATION3'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATIONQ'
       then TF.LaGrid.ColWidths[i]:=LgColCache;

      if (VALEUR(niveauT)>nivMax)
       then
        begin
         if TF.LaGrid.ColNames[i]='QBR_COEFFRETENU'
          then TF.LaGrid.ColWidths[i]:=LgColValeur;
         if TF.LaGrid.ColNames[i]='QBR_HISTO'
          then TF.LaGrid.ColWidths[i]:=LgColCache;
        end
       else
        begin
         if TF.LaGrid.ColNames[i]='QBR_COEFFRETENU'
          then TF.LaGrid.ColWidths[i]:=LgColCache;
         if TF.LaGrid.ColNames[i]='QBR_HISTO'
          then TF.LaGrid.ColWidths[i]:=LgColValeur;
        end;
    end;

 //CA prev
 if TabActive='TAB11'
  then
   for i:=1 to TF.LaGrid.colcount do
     begin
      if TF.LaGrid.ColNames[i]='QBR_VALEURAXE'
       then TF.LaGrid.ColWidths[i]:=LgColValeur;
      //CA1
      if TF.LaGrid.ColNames[i]='QBR_REF1'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OP1'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_REFPRCT1'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OPPRCT1'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      //CA2
      if TF.LaGrid.ColNames[i]='QBR_REF2'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OP2'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_REFPRCT2'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OPPRCT2'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      //CA3
      if TF.LaGrid.ColNames[i]='QBR_REF3'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OP3'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_REFPRCT3'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_OPPRCT3'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      //QTE
      if TF.LaGrid.ColNames[i]='QBR_QTEREF'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_QTEREFPRCT'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_QTEC'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_QTECPRCT'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      //coeff
      if TF.LaGrid.ColNames[i]='QBR_HISTO'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_REALISE'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_PREVU'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_COEFFCALCUL'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_COEFFRETENU'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_QTERETENUE'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_CARETENU'
       then TF.LaGrid.ColWidths[i]:=LgColValeur;
      if TF.LaGrid.ColNames[i]='QBR_HISTOCA'
       then TF.LaGrid.ColWidths[i]:=LgColValeur;
      if TF.LaGrid.ColNames[i]='QBR_PREVUCA'
       then TF.LaGrid.ColWidths[i]:=LgColValeur;
      if TF.LaGrid.ColNames[i]='QBR_REALISECA'
       then TF.LaGrid.ColWidths[i]:=LgColValeur;

      if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATION1'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATION2'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATION3'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
      if TF.LaGrid.ColNames[i]='QBR_PRCTVARIATIONQ'
       then TF.LaGrid.ColWidths[i]:=LgColCache;
    end;

end;

procedure TOM_QBPARBRE.Load(const edtSaisi,edtCorrectif,
          ChpSaisi,ChpCorrectif,ChpCorrectifPrct:hString;
          var EvolPrctSvg,SaisieValSvg,CorrectifSvg:hString);
begin
 SetControlText(edtSaisi,'');
 SetControlText(edtCorrectif,'');
 if getfield(ChpSaisi)<>0
  then SetControlText(edtSaisi,getfield(ChpSaisi));
 if getfield(ChpCorrectif)<>0
  then SetControlText(edtCorrectif,getfield(ChpCorrectif));

 EvolPrctSvg:=FloatToStr(THNumEdit(GetControl(ChpCorrectifPrct)).value);
 SaisieValSvg:=FloatToStr(THNumEdit(GetControl(ChpSaisi)).value);
 CorrectifSvg:=GetControlText(edtCorrectif);
end;

procedure TOM_QBPARBRE.OnLoadRecord ;
var i:integer;
begin
 Inherited ;
 OkModifValBloque:=false;

 if THCheckBox(GetControl('QBR_VALBLOQUE')).checked=true
  then
   begin
    //CA1
    THEdit(GetControl('QBR_EVOLPRCT1')).Font.Color:=clred;
    THEdit(GetControl('EDTEVOLVAL1')).Font.Color:=clred;
    THEdit(GetControl('EDTSAISI1')).Font.Color:=clred;
    //CA2
    THEdit(GetControl('QBR_EVOLPRCT2')).Font.Color:=clred;
    THEdit(GetControl('EDTEVOLVAL2')).Font.Color:=clred;
    THEdit(GetControl('EDTSAISI2')).Font.Color:=clred;
    //CA3
    THEdit(GetControl('QBR_EVOLPRCT3')).Font.Color:=clred;
    THEdit(GetControl('EDTEVOLVAL3')).Font.Color:=clred;
    THEdit(GetControl('EDTSAISI3')).Font.Color:=clred;
    //CA4
    THEdit(GetControl('QBR_EVOLPRCT4')).Font.Color:=clred;
    THEdit(GetControl('EDTEVOLVAL4')).Font.Color:=clred;
    THEdit(GetControl('EDTSAISI4')).Font.Color:=clred;
    //CA5
    THEdit(GetControl('QBR_EVOLPRCT5')).Font.Color:=clred;
    THEdit(GetControl('EDTEVOLVAL5')).Font.Color:=clred;
    THEdit(GetControl('EDTSAISI5')).Font.Color:=clred;
    //CA6
    THEdit(GetControl('QBR_EVOLPRCT6')).Font.Color:=clred;
    THEdit(GetControl('EDTEVOLVAL6')).Font.Color:=clred;
    THEdit(GetControl('EDTSAISI6')).Font.Color:=clred;
    //QTE
    THEdit(GetControl('QBR_EVOLQTEPRCT')).Font.Color:=clred;
    THEdit(GetControl('EDTEVOLVALQTE')).Font.Color:=clred;
    THEdit(GetControl('EDTSAISIQTE')).Font.Color:=clred;
    { EVI / Empêche la saisie si la case valeur bloquée est cochée }
    if ContextBP in [0,1,2] then THEdit(GetControl('PAGE')).Enabled:=false
    else
    begin
      THGroupBox(GetControl('BOX1')).Enabled:=false;
      THGroupBox(GetControl('BOX')).Enabled:=false;
      THGroupBox(GetControl('BOX3')).Enabled:=false;
      THGroupBox(GetControl('BOX4')).Enabled:=false;
      THGroupBox(GetControl('BOX5')).Enabled:=false;
      THGroupBox(GetControl('BOX6')).Enabled:=false;
      THGroupBox(GetControl('BOX7')).Enabled:=false;
    end;
   end
  else
   begin
    //CA1
    THEdit(GetControl('QBR_EVOLPRCT1')).Font.Color:=clBlack;
    THEdit(GetControl('EDTEVOLVAL1')).Font.Color:=clBlack;
    THEdit(GetControl('EDTSAISI1')).Font.Color:=clBlack;
    //CA2
    THEdit(GetControl('QBR_EVOLPRCT2')).Font.Color:=clBlack;
    THEdit(GetControl('EDTEVOLVAL2')).Font.Color:=clBlack;
    THEdit(GetControl('EDTSAISI2')).Font.Color:=clBlack;
    //CA3
    THEdit(GetControl('QBR_EVOLPRCT3')).Font.Color:=clBlack;
    THEdit(GetControl('EDTEVOLVAL3')).Font.Color:=clBlack;
    THEdit(GetControl('EDTSAISI3')).Font.Color:=clBlack;
    //CA4
    THEdit(GetControl('QBR_EVOLPRCT4')).Font.Color:=clBlack;
    THEdit(GetControl('EDTEVOLVAL4')).Font.Color:=clBlack;
    THEdit(GetControl('EDTSAISI4')).Font.Color:=clBlack;
    //CA5
    THEdit(GetControl('QBR_EVOLPRCT5')).Font.Color:=clBlack;
    THEdit(GetControl('EDTEVOLVAL5')).Font.Color:=clBlack;
    THEdit(GetControl('EDTSAISI5')).Font.Color:=clBlack;
    //CA6
    THEdit(GetControl('QBR_EVOLPRCT6')).Font.Color:=clBlack;
    THEdit(GetControl('EDTEVOLVAL6')).Font.Color:=clBlack;
    THEdit(GetControl('EDTSAISI6')).Font.Color:=clBlack;
    //QTE
    THEdit(GetControl('QBR_EVOLQTEPRCT')).Font.Color:=clBlack;
    THEdit(GetControl('EDTEVOLVALQTE')).Font.Color:=clBlack;
    THEdit(GetControl('EDTSAISIQTE')).Font.Color:=clBlack;
    { EVI / Autorise la saisie si la case valeur bloquée est cochée }
    if ContextBP in [0,1,2] then THEdit(GetControl('PAGE')).Enabled:=true
    else
    begin
      THGroupBox(GetControl('BOX1')).Enabled:=true;
      THGroupBox(GetControl('BOX')).Enabled:=true;
      THGroupBox(GetControl('BOX3')).Enabled:=true;
      THGroupBox(GetControl('BOX4')).Enabled:=true;
      THGroupBox(GetControl('BOX5')).Enabled:=true;
      THGroupBox(GetControl('BOX6')).Enabled:=true;
      THGroupBox(GetControl('BOX7')).Enabled:=true;
    end;
   end;

  OkDuplication:=false;
  if not okNew
   then
    begin
     //CA1
     THEdit(GetControl('QBR_REFPRCT1')).Enabled:=false;
     THEdit(GetControl('QBR_REF1')).Enabled:=false;
     //CA2
     THEdit(GetControl('QBR_REFPRCT2')).Enabled:=false;
     THEdit(GetControl('QBR_REF2')).Enabled:=false;
     //CA3
     THEdit(GetControl('QBR_REFPRCT3')).Enabled:=false;
     THEdit(GetControl('QBR_REF3')).Enabled:=false;
     //CA4
     THEdit(GetControl('QBR_REFPRCT4')).Enabled:=false;
     THEdit(GetControl('QBR_REF4')).Enabled:=false;
     //CA5
     THEdit(GetControl('QBR_REFPRCT5')).Enabled:=false;
     THEdit(GetControl('QBR_REF5')).Enabled:=false;
     //CA6
     THEdit(GetControl('QBR_REFPRCT6')).Enabled:=false;
     THEdit(GetControl('QBR_REF6')).Enabled:=false;
     //QTE
     THEdit(GetControl('QBR_QTEREFPRCT')).Enabled:=false;
     THEdit(GetControl('QBR_QTEREF')).Enabled:=false;

     THEdit(GetControl('QBR_VALEURAXE')).Enabled:=false;
    end;

  //CA1
  Load('EDTSAISI1','EDTEVOLVAL1','QBR_SAISI1','QBR_EVOLVAL1',
       'QBR_EVOLPRCT1',evolprctSvg,saivalSvg,evolvalSvg);

  //Qte
  Load('EDTSAISIQTE','EDTEVOLVALQTE','QBR_SAISIQTE','QBR_EVOLQTE',
       'QBR_EVOLQTEPRCT',EvolQtePrctSvg,SaisieQteValSvg,EvolQteValSvg);

  //CA2
  Load('EDTSAISI2','EDTEVOLVAL2','QBR_SAISI2','QBR_EVOLVAL2',
       'QBR_EVOLPRCT2',EvolCAPrct2Svg,SaisieCAVal2Svg,EvolCAVal2Svg);

  //CA3
  Load('EDTSAISI3','EDTEVOLVAL3','QBR_SAISI3','QBR_EVOLVAL3',
       'QBR_EVOLPRCT3',EvolCAPrct3Svg,SaisieCAVal3Svg,EvolCAVal3Svg);

  //CA4
  Load('EDTSAISI4','EDTEVOLVAL4','QBR_SAISI4','QBR_EVOLVAL4',
       'QBR_EVOLPRCT4',EvolCAPrct4Svg,SaisieCAVal4Svg,EvolCAVal4Svg);

  //CA5
  Load('EDTSAISI5','EDTEVOLVAL5','QBR_SAISI5','QBR_EVOLVAL5',
       'QBR_EVOLPRCT5',EvolCAPrct5Svg,SaisieCAVal5Svg,EvolCAVal5Svg);

  //CA6
  Load('EDTSAISI6','EDTEVOLVAL6','QBR_SAISI6','QBR_EVOLVAL6',
       'QBR_EVOLPRCT6',EvolCAPrct6Svg,SaisieCAVal6Svg,EvolCAVal6Svg);

  if ContextBP=3 then
  begin
    if TF.GetValue('QBR_CODEAXE')='DELAI' then
    begin
      THNumEdit(GetControl('ETPREF1')).Visible:=false;
      THNumEdit(GetControl('ETPOP1')).Visible:=false;
      THNumEdit(GetControl('QBR_EVOLPRCT1')).Enabled:=true;
      THEdit(GetControl('EDTEVOLVAL1')).Enabled:=true;
      THEdit(GetControl('EDTSAISI1')).Enabled:=true;
      THNumEdit(GetControl('ETPQTEREF')).Visible:=false;
      THNumEdit(GetControl('ETPQTEC')).Visible:=false;
      THNumEdit(GetControl('QBR_EVOLQTEPRCT')).Enabled:=true;
      THEdit(GetControl('EDTEVOLVALQTE')).Enabled:=true;
      THEdit(GetControl('EDTSAISIQTE')).Enabled:=true;
      for i:= 3 to 7 do
      begin
        THNumEdit(GetControl('ETPREF'+IntToStr(i-1))).Visible:=false;
        THNumEdit(GetControl('ETPOP'+IntToStr(i-1))).Visible:=false;
        THNumEdit(GetControl('QBR_EVOLPRCT'+IntToStr(i-1))).Enabled:=true;
        THEdit(GetControl('EDTEVOLVAL'+IntToStr(i-1))).Enabled:=true;
        THEdit(GetControl('EDTSAISI'+IntToStr(i-1))).Enabled:=true;
      end;
    end
    else
    begin
      if TabValues[1]='ETP' then
      begin
        THNumEdit(GetControl('ETPREF1')).Visible:=true;
        THNumEdit(GetControl('ETPREF1')).Value := THNumEdit(GetControl('QBR_REF1')).Value / NbIntervalle;
        THNumEdit(GetControl('ETPOP1')).Visible:=true;
        THNumEdit(GetControl('ETPOP1')).Value := THNumEdit(GetControl('QBR_OP1')).Value / NbIntervalle;
        THNumEdit(GetControl('QBR_EVOLPRCT1')).Enabled:=false;
        THEdit(GetControl('EDTEVOLVAL1')).Enabled:=false;
        THEdit(GetControl('EDTSAISI1')).Enabled:=false;
      end
      else
      begin
        THNumEdit(GetControl('ETPREF1')).Visible:=false;
        THNumEdit(GetControl('ETPOP1')).Visible:=false;
        THNumEdit(GetControl('QBR_EVOLPRCT1')).Enabled:=true;
        THEdit(GetControl('EDTEVOLVAL1')).Enabled:=true;
        THEdit(GetControl('EDTSAISI1')).Enabled:=true;
      end;
      if TabValues[2]='ETP' then
      begin
        THNumEdit(GetControl('ETPQTEREF')).Visible:=true;
        THNumEdit(GetControl('ETPQTEREF')).Value := THNumEdit(GetControl('QBR_QTEREF')).Value / NbIntervalle;
        THNumEdit(GetControl('ETPQTEC')).Visible:=true;
        THNumEdit(GetControl('ETPQTEC')).Value := THNumEdit(GetControl('QBR_QTEC')).Value / NbIntervalle;
        THNumEdit(GetControl('QBR_EVOLQTEPRCT')).Enabled:=false;
        THEdit(GetControl('EDTEVOLVALQTE')).Enabled:=false;
        THEdit(GetControl('EDTSAISIQTE')).Enabled:=false;
      end
      else
      begin
        THNumEdit(GetControl('ETPQTEREF')).Visible:=false;
        THNumEdit(GetControl('ETPQTEC')).Visible:=false;
        THNumEdit(GetControl('QBR_EVOLQTEPRCT')).Enabled:=true;
        THEdit(GetControl('EDTEVOLVALQTE')).Enabled:=true;
        THEdit(GetControl('EDTSAISIQTE')).Enabled:=true;
      end;
      for i:= 3 to 7 do
      begin
        if TabValues[i]='ETP' then
        begin
          THNumEdit(GetControl('ETPREF'+IntToStr(i-1))).Visible:=true;
          THNumEdit(GetControl('ETPREF'+IntToStr(i-1))).Value := THNumEdit(GetControl('QBR_REF'+IntToStr(i-1))).Value / NbIntervalle;
          THNumEdit(GetControl('ETPOP'+IntToStr(i-1))).Visible:=true;
          THNumEdit(GetControl('ETPOP'+IntToStr(i-1))).Value := THNumEdit(GetControl('QBR_OP'+IntToStr(i-1))).Value / NbIntervalle;
          THNumEdit(GetControl('QBR_EVOLPRCT'+IntToStr(i-1))).Enabled:=false;
          THEdit(GetControl('EDTEVOLVAL'+IntToStr(i-1))).Enabled:=false;
          THEdit(GetControl('EDTSAISI'+IntToStr(i-1))).Enabled:=false;
        end
        else
        begin
          THNumEdit(GetControl('ETPREF'+IntToStr(i-1))).Visible:=false;
          THNumEdit(GetControl('ETPOP'+IntToStr(i-1))).Visible:=false;
          THNumEdit(GetControl('QBR_EVOLPRCT'+IntToStr(i-1))).Enabled:=true;
          THEdit(GetControl('EDTEVOLVAL'+IntToStr(i-1))).Enabled:=true;
          THEdit(GetControl('EDTSAISI'+IntToStr(i-1))).Enabled:=true;
        end;
      end;
    end;
  end;


  { EVI / ORLI : Onglets CA & QTE Prevision }
  {
  //CA prev
  Load('EDTSAISIPREVCA','EDTEVOLVALPREVCA','QBR_SAISIPREVCA','QBR_EVOLPREVCA',
       'QBR_EVOLPREVCAPRCT',EvolCAPrctPrev,SaisieCAValPrevSvg,EvolCAValPrevSvg);

  //Qte prev
  Load('EDTSAISIPREVQTE','EDTEVOLVALPREVQTE','QBR_SAISIPREVQTE','QBR_EVOLPREVQTE',
       'QBR_EVOLPREVQTEPRC',EvolQtePrctPrev,SaisieQteValPrevSvg,EvolQteValPrevSvg);
  }
  
  //coeff
  CoeffRetenuSvg:=THNumEdit(GetControl('QBR_COEFFRETENU')).value;

  //-----------------> ORLI
  if BPOkOrli then AffichageColonneListe;
  //ORLI <-----------------

  if ContextBP=3 then TF. ForceShowCumulPanel := true ;
end ;

procedure TOM_QBPARBRE.OnChangeField ( F: TField ) ;
begin
  Inherited ;
  if F.FieldName = 'QBR_VALEURAXE'
   then
    begin
     THLabel(GetControl('LBLLIBVALEURAXE')).caption:=DonneLibelleValeurAxe(GetField('QBR_CODEAXE'),GetField('QBR_VALEURAXE'));
    end;
end ;

procedure TOM_QBPARBRE.OnArgument ( S: String ) ;
var i:integer;
codeSession:hString;
Q:TQuery;
DateDeb,DateFin:TDateTime;
begin
  Inherited ;
  if (ecran<>nil) and (ecran is TFSaisieList ) then TF := TFSaisieList(Ecran).LeFiltre
  else TF := nil;
  okNew:=false;
  codeSession:=TrouveArgument(S,'SESSION','');
  OkInitCoeff:=SessionInitCoeff(codeSession);
  OkInitPrev:=SessionInitPrev(codeSession);
  OkInitObj:=OkSessionObjectif(codeSession);
  OkInitNivTaille:=(SessionCalculParTaille(codeSession)) or (SessionEclateeParTaille(codeSession));
  nivMaxSession:=ChercheNivMaxSession(codeSession);

  Q:=OpenSQL('SELECT QBS_VALAFFH1,QBS_VALAFFH2,QBS_VALAFFH3,QBS_VALAFFH4,QBS_VALAFFH5,QBS_VALAFFH6,QBS_VALAFFH7'+
             ' FROM QBPSESSIONBP WHERE QBS_CODESESSION="'+CodeSession+'"',true);
  if not Q.Eof then
  begin
    for i:=1 to 7 do
    begin               //Tab1 = Val2, Tab2 = Val1, Tab3 = Val3...
      if i = 1 then TabValues[i+1] := Q.Fields[i-1].AsString
      else if i = 2 then TabValues[i-1] := Q.Fields[i-1].AsString
      else TabValues[i] := Q.Fields[i-1].AsString
    end;
  end;
  Ferme(Q);

  if ContextBP=3 then
  begin
    dateDebFinSesssion(codeSession,DateDeb,DateFin);
    NbIntervalle:=DonneNbIntervalleMailleDateDebDateFin('4',DateDeb,DateFin);
  end;
end ;

procedure TOM_QBPARBRE.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_QBPARBRE.OnCancelRecord ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOM_QBPARBRE ] ) ;
end.
