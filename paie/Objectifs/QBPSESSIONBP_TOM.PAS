{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 15/12/2005
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : QBPSESSIONBP (QBPSESSIONBP)
Modifié le ... : 30/10/2006
Description .. : Gestion de l'entrée par la fiche GCTYPEDA_FIL
Modifié le ... : 26/01/2007 FQ N° 14682
Description .. : Passage de la valeur analysée et du 1 er axe en lecture seule
Suite......... : dans le cas ou l'on vient des parametres budgétaires sur DA
Mots clefs ... : TOM;QBPSESSIONBP
*****************************************************************}
Unit QBPSESSIONBP_TOM ;

Interface

Uses UTOM,HEnt1;

Type
  TOM_QBPSESSIONBP = Class (TOM)
  public
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnArgument ( StArgument: String )   ; override ;
  private
    FTypeAction : TActionFiche;
    { EVI / OnNewRecord entraine OnLoadRecord --> Problème en création }
    Loaded,OkToClose,TabValAffSelected : boolean;
    sArg, sArguments : string;    //CEGID-CCMX FQ N° 14840
    procedure BTNINITG_OnClick(Sender: TObject);
    procedure POPLOI_OnPopup(Sender: TObject);
    procedure ajoutloi_OnClick(Sender: TObject);
    procedure calculloi_OnClick(Sender: TObject);
    procedure POPSESSION_OnPopup(Sender: TObject);
    procedure ControleS_OnClick(Sender: TObject);
    procedure Sdetaillee_OnClick(Sender: TObject);
    procedure POPINITDELAI_OnPopup(Sender: TObject);
    procedure Mois445_OnClick(Sender: TObject) ;
    procedure Semaine_OnClick(Sender: TObject);
    procedure Quinzaine_OnClick(Sender: TObject);
    procedure Mois_OnClick(Sender: TObject);
    procedure Trimestre_OnClick(Sender: TObject);
    procedure Quadrimestre_OnClick(Sender: TObject);
    procedure AxeAnalyse_OnChange(Sender: TObject);
    procedure AxeAnalyse_OnEnter(Sender: TObject);
    procedure AxeAnalyse_OnLoad(Sender: TObject);
    procedure RADIOGRP_OnClick(Sender: TObject);
    procedure JOURD1_OnChange(Sender: TObject);
    procedure AxeRestriction_OnEnter(Sender: TObject);
    procedure AxeRestriction_OnChange(Sender: TObject);
    procedure ValideS_OnClick (Sender: TObject);
    procedure InvalideS_OnClick (Sender: TObject);
    procedure BTNDUPLIQUER_OnClick(Sender: TObject);
    procedure ValeurAxeRestriction_OnClick(Sender: TObject);
    function FuncBPSessionMAJChamp : boolean;
    procedure ValAff_OnChange(Sender: TObject);
    procedure ValAff_OnEnter(Sender: TObject);
    procedure ValRub_OnChange(Sender: TObject);
    procedure ValRub_OnEnter(Sender: TObject);
    procedure PageChange(Sender: TObject);
    procedure TraitementRadioGroup;
    {$IFDEF PAIEGRH}
    procedure LibellesLibrePaie(Champ:string;NbChamp:integer);
    procedure LibellesRubSalPaie(Champ:string);
    {$ENDIF PAIEGRH}
    procedure MultiSocS_OnClick(Sender: TObject);
  end ;

Implementation
//CEGID-CCMX le 30/10/2006 DEBUT

{ EVI / ToDo : [Voir avec L.Meunier] Fonction GetArgumentString = wCommuns (Mode) & UtilPGI (GC) }
Uses StdCtrls,Classes,HMsgBox,sysutils,
     HDB,paramsoc,HTB97,HCtrls,Controls,StrUtils,
     {$IFDEF EAGLCLIENT}MaineAGL,UtileAGL,eFiche,UTob,
     {$ELSE}FE_Main,Fiche,DBCtrls,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     {$ENDIF}
     {$IFDEF MODE} wCommuns, {$ELSE} UtilPGI, {$ENDIF}
     {$IFDEF PAIEGRH}EntPaie,{$ENDIF PAIEGRH}
     SynScriptBP,BPBasic,BPFctSession,Uutil,CstCommun,BPUtil;

var
  sOrigine : string;
  LastValue,ChampValAffOK:string;
//CEGID-CCMX le 30/10/2006 FIN

procedure TOM_QBPSESSIONBP.OnNewRecord ;
var TabCodeAxe : hString;
  sValeur : string;  //CEGID-CCMX le 12/04/2007 FQ N° 14840
begin
  Inherited ;
  if Ecran.Name='QUFPBPSESSIONENT' then
  begin
    SetControlText('jourd2', FuncSynBPDateDebutAnneeMoinsUn([0],0));
    SetControlText('jourd2_', FuncSynBPDateFinAnneeMoinsUn([0],0));
    SetControlText('jourd1_', FuncSynBPDateFinAnnee([0],0));
    SetControlText('jourd1', FuncSynBPDateDebutAnnee([0],0));

    if GetField('QBS_DATEVALIDATION') = idate1900
    then SetControlVisible('QBS_DATEVALIDATION', false)
    else SetControlVisible('QBS_DATEVALIDATION', true);
    //en mode nouveau
    //on a accès au choix de la méthode d'initialisation

    if ContextBP=3 then SetControlEnabled('radiogrp', FALSE)
    else SetControlEnabled('radiogrp', TRUE);
    //on n'a pas accès aux differnts boutons

    SetControlEnabled('btnloi',FALSE);
    SetControlEnabled('btninit',FALSE);
    SetControlVisible('btninitg',FALSE);
    //FQ Mode 15089
    SetControlEnabled('btnmenu',FALSE);

    //on a accès au bouton dupliquer
    SetControlVisible('btndupliquer',TRUE);

    SetControlProperty('radiogrp','ITEMINDEX',0);

    //mise à jour noms des tablettes codes axes et valeurs axes
    TabCodeAxe:=FuncSynBPTabletteCodeAxe([0],0);

    SetControlProperty('COMBOBOXAXE1','DATATYPE',TabCodeAxe);
    SetControlProperty('COMBOBOXAXE2','DATATYPE',TabCodeAxe);
    SetControlProperty('COMBOBOXAXE3','DATATYPE',TabCodeAxe);
    SetControlProperty('COMBOBOXAXE4','DATATYPE',TabCodeAxe);
    SetControlProperty('COMBOBOXAXE5','DATATYPE',TabCodeAxe);
    SetControlProperty('COMBOBOXAXE6','DATATYPE',TabCodeAxe);
    SetControlProperty('COMBOBOXAXE7','DATATYPE',TabCodeAxe);
    SetControlProperty('COMBOBOXAXE8','DATATYPE',TabCodeAxe);
    SetControlProperty('COMBOBOXAXE9','DATATYPE',TabCodeAxe);
    SetControlProperty('COMBOBOXAXE10','DATATYPE',TabCodeAxe);

    {$IFDEF PAIEGRH}
    if ContextBP = 3 then LibellesLibrePaie('COMBOBOXAXE',10);
    {$ENDIF PAIEGRH}

    SetControlProperty('QBS_CODEAXE1','DATATYPE',TabCodeAxe);
    SetControlProperty('QBS_CODEAXE2','DATATYPE',TabCodeAxe);
    SetControlProperty('QBS_CODEAXE3','DATATYPE',TabCodeAxe);
    SetControlProperty('QBS_CODEAXE4','DATATYPE',TabCodeAxe);

    {$IFDEF PAIEGRH}
    if ContextBP = 3 then LibellesLibrePaie('QBS_CODEAXE',4);
    {$ENDIF PAIEGRH}

    SetControlText('TabPlus', BPTablettePlus('0'));
    //on initialise des valeurs

    SetField('QBS_TYPENATURE','1');

    //CEGID-CCMX le 30/10/2006 : on force la valeur affichée au CA_HT Net
    if (sOrigine = 'TYPEDA') then
    begin
      //CEGID-CCMX le 12/04/2007 FQ N° 14840 DEBUT
      if (sArg = 'CREATION') then
      begin
        sValeur := GetArgumentvalue(SArguments, 'LIBELLE');
        SetField('QBS_LIBSESSION',sValeur);
        sValeur := GetArgumentvalue(SArguments, 'DEBUT');
        SetField('JOURD1',sValeur);
        sValeur := GetArgumentvalue(SArguments, 'FIN');
        SetControlText('JOURD1_',sValeur);
        sValeur := GetArgumentvalue(SArguments, 'DEBUTANT');
        SetControlText('JOURD2',sValeur);
        sValeur := GetArgumentvalue(SArguments, 'FINANT');
        SetControlText('JOURD2_',sValeur);
      SetField('QBS_VUEARBRE','PHT');
      TabCodeAxe := 'QUTBPAXETYPEDA';
//      SetField('QBS_CODEAXES1','046');  //CEGID-CCMX le 26/01/2007 FQ N° 14682
      SetControlText('COMBOBOXAXE1','046'); { EVI : Correction pour TypeDA avec axes = THValComboBox }
      SetControlText('COMBOBOXAXE2','003'); { EVI : Correction pour TypeDA avec axes = THValComboBox }
      SetControlText('COMBOBOXAXE3','047'); { EVI : Correction pour TypeDA avec axes = THValComboBox }

//        SetField('QBS_CODEAXES2','003');
//        SetField('QBS_CODEAXES3','047');
      end;
      //CEGID-CCMX le 12/04/2007 FQ N° 14840 FIN
    end
    else
    begin
      Case ContextBP of
        0,1 : SetField('QBS_VUEARBRE','PTC');
        2 : SetField('QBS_VUEARBRE','CD1');
      end
    end;

    SetField('QBS_NATURECMD',TraduireMemoire('<<Tous>>'));
    SetField('QBS_NATURE2',TraduireMemoire('<<Tous>>'));

    SetField('QBS_SESSIONVALIDE','-');
  end;
end ;

procedure TOM_QBPSESSIONBP.OnDeleteRecord ;
{$IFNDEF MODE}{$IFDEF GCGC}
var sSessionN, sSessionN1, sCodeSession : string; {$ENDIF}{$ENDIF}
begin
  Inherited ;
  //CEGID-CCMX le 16/11/2006 : si c'est une session liée aux DA, suppression impossible
  if Ecran.Name='QUFPBPSESSIONENT' then
  begin
    {$IFNDEF MODE}
    {$IFDEF GCGC}
    if (sOrigine = 'TYPEDA') then
    begin
      sSessionN := GetParamsocSecur('SO_GCCODESESSIONN','');
      sSessionN1 := GetParamsocSecur('SO_GCCODESESSIONN1','');
       // le 16/01/2007 : en fait on limite pas seulement au session N ou N+1 mais aux
       // sessions utilisées
      sCodeSession := GetField('QBS_CODESESSION');
      if (sSessionN = GetField('QBS_CODESESSION')) or ( sSessionN1 = GetField('QBS_CODESESSION'))
      or (ExisteSQL('SELECT DA_CODESESSION FROM PIECEDA WHERE DA_CODESESSION="' + sCodeSession +'"'))
      then
      begin
        LastError := 1;
        LastErrorMsg := 'Vous ne pouvez pas supprimer une session attachée aux budgets pour les DA';
      end;
    end
    else
    {$ENDIF}
    {$ENDIF}
    begin
      //ExecuteSQL('DELETE FROM QBPSESSIONBP WHERE QBS_CODESESSION="'+GetControlText('QBS_CODESESSION')+'"');
      ExecuteSQL('DELETE FROM QBPARBRE WHERE QBR_CODESESSION="'+GetControlText('QBS_CODESESSION')+'"');
      ExecuteSQL('DELETE FROM QBPARBREDETAIL WHERE QBH_CODESESSION="'+GetControlText('QBS_CODESESSION')+'"');
      ExecuteSQL('DELETE FROM QBPLOI WHERE QBO_CODESESSION="'+GetControlText('QBS_CODESESSION')+'"');
      ExecuteSQL('DELETE FROM QBPCUBETMP WHERE QBQ_CODESESSION="'+GetControlText('QBS_CODESESSION')+'"');
    end;
  end ;
end ;

procedure TOM_QBPSESSIONBP.OnUpdateRecord ;
var i,j,LastAxe:integer;
    resAnalyse, resA1, resA2, resA3, resA4, resA5, resA6, resA7, resA8, resA9, resA10 : boolean;
    resRestrict, resRestrictValeur, resR1, resR2, resR3, resR4 : boolean;
    resValAff, resDate, resD1, resD2, resSession, resOrdre : boolean;
    resDouble,resSalarie : boolean;
    NbresValAff : integer;
    CurrentAxe:string;
  bAxeTypeDA : boolean; //CCMX-CEGID le 30/10/2006
begin
  Inherited ;
  if Ecran.Name='QUFPBPSESSIONENT' then
  begin
    FuncBPSessionMAJChamp ;
    LastAxe:=0;
    resValAff := true;
    resSession:=true ;
    resDouble:=true;
    resSalarie:=true;
    resRestrictValeur:=true;

    if TFFiche(Ecran).FTypeAction = taCreat
    then resSession:=FuncSynBPValeurUnique([GetControlText('QBS_CODESESSION'),'QBPSESSIONBP','QBS_CODESESSION','session'],0);
    //CEGID-CCMX le 30/10/2006 : si c'est une session liée aux DA, valeur affichee=CA HT Net
    if (sOrigine = 'TYPEDA') then
    begin
      if(GetControlText('QBS_VUEARBRE') <> 'PHT') then
      begin
        PGIInfo('Les sessions d''objectifs liées aux demandes d''achats ne peuvent retourner que le CA HT Net');
        SetField('QBS_VUEARBRE','PHT');
      end;
      //CEGID-CCMX le 16/11/2006 DEBUT
      if (GetControlText('QBS_CODEAXES1') <> '046') then
      begin
        PGIInfo('Les sessions d''objectifs liées aux demandes d''achats doivent avoir les types de DA en premier axe d''analyse');
        SetField('QBS_CODEAXES1','046');
      end;
      //CEGID-CCMX le 16/11/2006 FIN
    end
    else
    begin
      bAxeTypeDA := False;
      if ( GetControlText('QBS_CODEAXES1') = '046') or (GetControlText('QBS_CODEAXES1') = '047') then
          bAxeTypeDA := True;
      if ( GetControlText('QBS_CODEAXES2') = '046') or (GetControlText('QBS_CODEAXES2') = '047') then
          bAxeTypeDA := True;
      if ( GetControlText('QBS_CODEAXES3') = '046') or (GetControlText('QBS_CODEAXES3') = '047') then
          bAxeTypeDA := True;
      if ( GetControlText('QBS_CODEAXES4') = '046') or (GetControlText('QBS_CODEAXES4') = '047') then
          bAxeTypeDA := True;
      if ( GetControlText('QBS_CODEAXES5') = '046') or (GetControlText('QBS_CODEAXES5') = '047') then
          bAxeTypeDA := True;
      if ( GetControlText('QBS_CODEAXES6') = '046') or (GetControlText('QBS_CODEAXES6') = '047') then
          bAxeTypeDA := True;
      if ( GetControlText('QBS_CODEAXES7') = '046') or (GetControlText('QBS_CODEAXES7') = '047') then
          bAxeTypeDA := True;
      if ( GetControlText('QBS_CODEAXES8') = '046') or (GetControlText('QBS_CODEAXES8') = '047') then
          bAxeTypeDA := True;
      if ( GetControlText('QBS_CODEAXES9') = '046') or (GetControlText('QBS_CODEAXES9') = '047') then
          bAxeTypeDA := True;
      if ( GetControlText('QBS_CODEAXES10') = '046') or (GetControlText('QBS_CODEAXES10') = '047') then
          bAxeTypeDA := True;
      if (bAxeTypeDA = True) then
        PGIInfo('Attention Les axes 046 et 047 ne sont possibles que pour les objectifs liées aux demandes d''achats');
    end;
    //CEGID-CCMX le 30/10/2006 Fin

    //controle sur les 10 axes analyse
    resA1:=ControleAxe(GetControlText('QBS_CODEAXES1'),'','Axe 1');
    resA2:=ControleAxe(GetControlText('QBS_CODEAXES2'),'','Axe 2');
    resA3:=ControleAxe(GetControlText('QBS_CODEAXES3'),'','Axe 3');
    resA4:=ControleAxe(GetControlText('QBS_CODEAXES4'),'','Axe 4');
    resA5:=ControleAxe(GetControlText('QBS_CODEAXES5'),'','Axe 5');
    resA6:=ControleAxe(GetControlText('QBS_CODEAXES6'),'','Axe 6');
    resA7:=ControleAxe(GetControlText('QBS_CODEAXES7'),'','Axe 7');
    resA8:=ControleAxe(GetControlText('QBS_CODEAXES8'),'','Axe 8');
    resA9:=ControleAxe(GetControlText('QBS_CODEAXES9'),'','Axe 9');
    resA10:=ControleAxe(GetControlText('QBS_CODEAXES10'),'','Axe 10');

    //controle sur les 4 axes de restriction
    resR1:=ControleAxe(GetControlText('QBS_CODEAXE1'),'','Restriction 1');
    resR2:=ControleAxe(GetControlText('QBS_CODEAXE2'),'','Restriction 2');
    resR3:=ControleAxe(GetControlText('QBS_CODEAXE3'),'','Restriction 3');
    resR4:=ControleAxe(GetControlText('QBS_CODEAXE4'),'','Restriction 4');

    //controle sur les 4 groupes de valeur des axes de restriction
    for i:=1 to 4 do
    begin
      if ((GetControlText('QBS_CODEAXE'+IntToStr(i))<>'') and (GetControlText('QBS_VALEURAXE'+IntToStr(i))='')) then
      begin
        LastError:=1;
        LastErrorMsg:=TraduireMemoire('La valeur pour l''axe de restriction '+IntToStr(i)+' est vide');
        resRestrictValeur:=false;
      end;
    end;

    resD1:=STRTODATE(GetControlText('JOURD1'))<STRTODATE(GetControlText('JOURD1_'));
    resD2:=STRTODATE(GetControlText('JOURD2'))<STRTODATE(GetControlText('JOURD2_'));

    if (resD1=false) or (resD2=false)
    then
    begin
      LastError:=1;
      LastErrorMsg:=TraduireMemoire('La date de fin doit être postérieure à la date de début.');
    end;

    if ContextBP = 3 then
    begin
      NbresValAff := 0;
      for i:=1 to ThSpinEdit(GetControl('QBS_NBVALAFF')).Value do
      begin
        if GetControlText('QBS_VALAFFLIB'+IntToStr(i)) = '' then NbresValAff := i;
        if NbresValAff <> 0 then break;
        if GetControlText('QBS_VALAFFH'+IntToStr(i)) = '' then NbresValAff := i;
        if NbresValAff <> 0 then break;
        if ((GetControlVisible('QBS_VALAFFH'+IntToStr(i)+'A')) AND (GetControlText('QBS_VALAFFH'+IntToStr(i)+'A') = '')) then NbresValAff := i;
        if NbresValAff <> 0 then break;
        if ((GetControlVisible('QBS_VALAFFH'+IntToStr(i)+'B')) AND (GetControlText('QBS_VALAFFH'+IntToStr(i)+'B') = '')) then NbresValAff := i;
        if NbresValAff <> 0 then break;
        if GetControlText('QBS_VALAFFR'+IntToStr(i)) = '' then NbresValAff := i;
        if NbresValAff <> 0 then break;
        if ((GetControlVisible('QBS_VALAFFR'+IntToStr(i)+'A')) AND (GetControlText('QBS_VALAFFR'+IntToStr(i)+'A') = '')) then NbresValAff := i;
        if NbresValAff <> 0 then break;
        if ((GetControlVisible('QBS_VALAFFR'+IntToStr(i)+'B')) AND (GetControlText('QBS_VALAFFR'+IntToStr(i)+'B') = '')) then NbresValAff := i;
        if NbresValAff <> 0 then break;
      end;
      if NbresValAff <> 0 then
      begin
        LastError:=1;
        LastErrorMsg:=TraduireMemoire('Les champs pour la valeur affichée '+IntToStr(i)+' ne sont pas correctement renseignés.');
        resValAff := false;
      end;
    end;

    resAnalyse := (resA1 and resA2 and resA3 and resA4 and resA5 and resA6 and resA7 and resA8 and resA9 and resA10);
    resRestrict := (resR1 and resR2 and resR3 and resR4 and resRestrictValeur);
    resDate := (resD1 and resD2);

    resOrdre:=true;

    for i := 10 downto 1 do
    begin
      if GetControlText('QBS_CODEAXES'+IntToStr(i))<>'' then
      begin
        LastAxe := i;
        break;
      end;
    end;

    if LastAxe <> 0 then
    begin
      for i := LastAxe downto 1 do
      begin
        if GetControlText('QBS_CODEAXES'+IntToStr(i))='' then resOrdre:=false;
      end;
      if resOrdre=false then
      begin
        LastError:=1;
        LastErrorMsg:=TraduireMemoire('Les axes d''analyse doivent être renseignés dans l''ordre.');
      end;
    end
    else
    begin
      resOrdre:=false;
      LastError:=1;
      LastErrorMsg:=TraduireMemoire('Vous devez renseigner les axes.');
    end;

    //Contrôle axe Salarié présent en dernière position + redondance d'axes.
    if ContextBP = 3 then
    begin
      i:=1;
      While GetControlText('QBS_CODEAXES'+IntToStr(i))<>'' do
      begin
        CurrentAxe:=GetControlText('QBS_CODEAXES'+IntToStr(i));
        for j:=i+1 to 10 do
        begin
          if CurrentAxe=GetControlText('QBS_CODEAXES'+IntToStr(j)) then
          begin
            resDouble:=false;
            LastError:=1;
            LastErrorMsg:=TraduireMemoire('Des axes d''analyse sont en doublon.');
          end;
        end;
        i:=i+1;
      end;

      if GetControlText('QBS_CODEAXES'+IntToStr(i-1))<>'011' then
      begin
        resSalarie:=false;
        LastError:=1;
        LastErrorMsg:=TraduireMemoire('Les sessions d''objectifs paie doivent avoir l''axe salarié en dernier axe d''analyse.');
      end;
    end;


    OkToClose := (resSession and resAnalyse and resRestrict and resDate and resOrdre and resValAff and resDouble and resSalarie );


    if OkToClose=true
    then
    begin
      if GetControlText('edtokloi')='6'
      then ExecuteSql('insert into qbploi (QBO_BPM1,QBO_BPM10,QBO_BPM11,QBO_BPM12,'+
          'QBO_BPM13,QBO_BPM14,QBO_BPM15,QBO_BPM16,QBO_BPM17,QBO_BPM18,QBO_BPM19,'+
          'QBO_BPM2,QBO_BPM20,QBO_BPM21,QBO_BPM22,QBO_BPM23,QBO_BPM24,QBO_BPM25,'+
          'QBO_BPM26,QBO_BPM27,QBO_BPM28,QBO_BPM29,QBO_BPM3,QBO_BPM30,QBO_BPM31,'+
          'QBO_BPM32,QBO_BPM33,QBO_BPM34,QBO_BPM35,QBO_BPM36,QBO_BPM37,QBO_BPM38,'+
          'QBO_BPM39,QBO_BPM4,QBO_BPM40,QBO_BPM41,QBO_BPM42,QBO_BPM43,QBO_BPM44,'+
          'QBO_BPM45,QBO_BPM46,QBO_BPM47,QBO_BPM48,QBO_BPM49,QBO_BPM5,QBO_BPM50,'+
          'QBO_BPM51,QBO_BPM52,QBO_BPM6,QBO_BPM7,QBO_BPM8,QBO_BPM9,QBO_CODEAXER1,'+
          'QBO_CODEAXER10,QBO_CODEAXER2,QBO_CODEAXER3,QBO_CODEAXER4,QBO_CODEAXER5,'+
          'QBO_CODEAXER6,QBO_CODEAXER7,QBO_CODEAXER8,QBO_CODEAXER9,QBO_CODELOI,'+
          'QBO_CODESESSION,QBO_CODESTRUCT,QBO_CREATEUR,QBO_DATECREATION,'+
          'QBO_DATEDEBECLAT,QBO_DATEFINECLAT,QBO_DATEMODIF,QBO_LIBLOI,QBO_MAILLE,'+
          'QBO_MODEECLAT,QBO_NUMMAILLE,QBO_PLAGE,QBO_UTILISATEUR,QBO_VALAXER1,'+
          'QBO_VALAXER10,QBO_VALAXER2,QBO_VALAXER3,QBO_VALAXER4,QBO_VALAXER5,'+
          'QBO_VALAXER6,QBO_VALAXER7,QBO_VALAXER8,QBO_VALAXER9) select '+
          'QBO_BPM1,QBO_BPM10,QBO_BPM11,QBO_BPM12,'+
          'QBO_BPM13,QBO_BPM14,QBO_BPM15,QBO_BPM16,QBO_BPM17,QBO_BPM18,QBO_BPM19,'+
          'QBO_BPM2,QBO_BPM20,QBO_BPM21,QBO_BPM22,QBO_BPM23,QBO_BPM24,QBO_BPM25,'+
          'QBO_BPM26,QBO_BPM27,QBO_BPM28,QBO_BPM29,QBO_BPM3,QBO_BPM30,QBO_BPM31,'+
          'QBO_BPM32,QBO_BPM33,QBO_BPM34,QBO_BPM35,QBO_BPM36,QBO_BPM37,QBO_BPM38,'+
          'QBO_BPM39,QBO_BPM4,QBO_BPM40,QBO_BPM41,QBO_BPM42,QBO_BPM43,QBO_BPM44,'+
          'QBO_BPM45,QBO_BPM46,QBO_BPM47,QBO_BPM48,QBO_BPM49,QBO_BPM5,QBO_BPM50,'+
          'QBO_BPM51,QBO_BPM52,QBO_BPM6,QBO_BPM7,QBO_BPM8,QBO_BPM9,QBO_CODEAXER1,'+
          'QBO_CODEAXER10,QBO_CODEAXER2,QBO_CODEAXER3,QBO_CODEAXER4,QBO_CODEAXER5,'+
          'QBO_CODEAXER6,QBO_CODEAXER7,QBO_CODEAXER8,QBO_CODEAXER9,QBO_CODELOI,'+
          '"'+GetControlText('QBS_CODESESSION')+'",QBO_CODESTRUCT,QBO_CREATEUR,QBO_DATECREATION,'+
          'QBO_DATEDEBECLAT,QBO_DATEFINECLAT,QBO_DATEMODIF,QBO_LIBLOI,QBO_MAILLE,'+
          'QBO_MODEECLAT,QBO_NUMMAILLE,QBO_PLAGE,QBO_UTILISATEUR,QBO_VALAXER1,'+
          'QBO_VALAXER10,QBO_VALAXER2,QBO_VALAXER3,QBO_VALAXER4,QBO_VALAXER5,'+
          'QBO_VALAXER6,QBO_VALAXER7,QBO_VALAXER8,QBO_VALAXER9 '+
          'from qbploi where qbo_codesession="'+GetControlText('SessionDup')+'"');
      if GetControlText('OkDupTotale')='OUI'
      then
      begin
        DuplicationArbreBP(GetControlText('SessionDup'),GetControlText('QBS_CODESESSION'));
      end ;
      if GetControlText('NivDupPartielle')<>''
      then
      begin
        DuplicationArbrePartielleBP(GetControlText('SessionDup'),GetControlText('QBS_CODESESSION'),GetControlText('NivDupPartielle'));
      end ;
    end ;
  end ;
end;

procedure TOM_QBPSESSIONBP.OnAfterUpdateRecord ;
begin
  Inherited ;
  if Ecran.Name='QUFPBPSESSIONENT' then
  begin
    if TFFiche(Ecran).FTypeAction = taCreat then TFFiche(Ecran).FTypeAction := taModif;
    { EVI / Bloque la possibilité de changer la valeur à afficher une fois l'initialisation faite }
    SetControlEnabled('QBS_VUEARBRE', FALSE);
    SetControlEnabled('QBS_CODESESSION', FALSE);
    SetControlEnabled('radiogrp',FALSE);
    TraitementRadioGroup;
    ProcSynBPMAJDatesLoi([GetField('QBS_CODESESSION'),GetField('QBS_DATEDEBC'),GetField('QBS_DATEFINC')],0);
    //CEGID-CCMX le 12/04/2007 FQ N°14840 DEBUT
    if (sOrigine = 'TYPEDA') then
    begin
      if(GetArgumentvalue(SArguments, 'ANNEE')='N') then
        SetParamsoc('SO_GCCODESESSIONN',GetField('QBS_CODESESSION'));
      if(GetArgumentvalue(SArguments, 'ANNEE')='N1') then
        SetParamsoc('SO_GCCODESESSIONN1',GetField('QBS_CODESESSION'));
    end;
    //CEGID-CCMX le 12/04/2007 FQ N°14840 FIN
  end;
end;

procedure TOM_QBPSESSIONBP.OnLoadRecord ;
var TabCodeAxe : hString;
{$IFDEF PAIEGRH}i:integer;{$ENDIF PAIEGRH}
begin
  Inherited ;
  if Ecran.Name='QUFPBPSESSIONENT' then
  begin
    if((FTypeAction <> taCreat) AND (Loaded = false))
    then
    begin
      //en mode load on n'a pas accès au bouton dupliquer
      SetControlEnabled('btndupliquer', false);
      if GetField('QBS_DATEVALIDATION') = idate1900 then SetControlVisible('QBS_DATEVALIDATION', false)
      else SetControlVisible('QBS_DATEVALIDATION', true);

      SetControlText('JOURD1', DateToStr(strToDateTime0(GetControlText('QBS_DATEDEBC'))));
      SetControlText('JOURD1_', DateToStr(strToDateTime0(GetControlText('QBS_DATEFINC'))));
      SetControlText('JOURD2', DateToStr(strToDateTime0(GetControlText('QBS_DATEDEBREF'))));
      SetControlText('JOURD2_', DateToStr(strToDateTime0(GetControlText('QBS_DATEFINREF'))));
      Loaded := true;

      //mise à jour noms des tablettes codes axes et valeurs axes
      //CCMX-CEGID : on filtre la liste des axes
      if (sOrigine = 'TYPEDA') then
       TabCodeAxe := 'QUTBPAXETYPEDA'
      else //CCMX-CEGID FIN
      TabCodeAxe:=FuncSynBPTabletteCodeAxe([0],0);
      SetControlText('TabPlus', BPTablettePlus('0'));

      SetControlProperty('COMBOBOXAXE1','DATATYPE',TabCodeAxe);
      SetControlProperty('COMBOBOXAXE2','DATATYPE',TabCodeAxe);
      SetControlProperty('COMBOBOXAXE3','DATATYPE',TabCodeAxe);
      SetControlProperty('COMBOBOXAXE4','DATATYPE',TabCodeAxe);
      SetControlProperty('COMBOBOXAXE5','DATATYPE',TabCodeAxe);
      SetControlProperty('COMBOBOXAXE6','DATATYPE',TabCodeAxe);
      SetControlProperty('COMBOBOXAXE7','DATATYPE',TabCodeAxe);
      SetControlProperty('COMBOBOXAXE8','DATATYPE',TabCodeAxe);
      SetControlProperty('COMBOBOXAXE9','DATATYPE',TabCodeAxe);
      SetControlProperty('COMBOBOXAXE10','DATATYPE',TabCodeAxe);

      {$IFNDEF PAIEGRH}
      SetControlText('COMBOBOXAXE1',GetControlText('QBS_CODEAXES1'));
      SetControlText('COMBOBOXAXE2',GetControlText('QBS_CODEAXES2'));
      SetControlText('COMBOBOXAXE3',GetControlText('QBS_CODEAXES3'));
      SetControlText('COMBOBOXAXE4',GetControlText('QBS_CODEAXES4'));
      SetControlText('COMBOBOXAXE5',GetControlText('QBS_CODEAXES5'));
      SetControlText('COMBOBOXAXE6',GetControlText('QBS_CODEAXES6'));
      SetControlText('COMBOBOXAXE7',GetControlText('QBS_CODEAXES7'));
      SetControlText('COMBOBOXAXE8',GetControlText('QBS_CODEAXES8'));
      SetControlText('COMBOBOXAXE9',GetControlText('QBS_CODEAXES9'));
      SetControlText('COMBOBOXAXE10',GetControlText('QBS_CODEAXES10'));
      {$ELSE}
      if ContextBP = 3 then LibellesLibrePaie('COMBOBOXAXE',10);
      if ContextBP = 3 then
      begin
        for i := 1 to 10 do
        SetControlText('COMBOBOXAXE'+IntToStr(i),GetControlText('QBS_CODEAXES'+IntToStr(i)))
      end;
      {$ENDIF PAIEGRH}

      SetControlProperty('QBS_CODEAXE1','DATATYPE',TabCodeAxe);
      SetControlProperty('QBS_CODEAXE2','DATATYPE',TabCodeAxe);
      SetControlProperty('QBS_CODEAXE3','DATATYPE',TabCodeAxe);
      SetControlProperty('QBS_CODEAXE4','DATATYPE',TabCodeAxe);

      {$IFDEF PAIEGRH}
      if ContextBP = 3 then LibellesLibrePaie('QBS_CODEAXE',4);
      {$ENDIF PAIEGRH}

      if GetControlText('QBS_METHODE') ='1'
      then SetControlProperty('radiogrp','ITEMINDEX',0)
      else SetControlProperty('radiogrp','ITEMINDEX',1);
      //on n'a pas accès au choix de la méthode d'initialisation

      SetControlEnabled('radiogrp',FALSE);
     { EVI / Bloque la possibilité de changer la valeur à affciher une fois l'initialisation faite }
      SetControlEnabled('QBS_VUEARBRE', FALSE);
      //suivant la methode affiche le bon choix dans le radiogrp
      //et rend accessible ou non les boutons
      TraitementRadioGroup;

      //si on est en consult on ne doit avoir accès aux boutons de calculs
      if (FTypeAction=TaConsult) //or (GetControlText('edtnew')='NEW')
      then
      begin
        SetControlEnabled('btninit', false);
        SetControlEnabled('btninitg', false);
        SetControlEnabled('btnloi', false);
      end ;

      //affichage du libelle de l'etat de la session init
      THCheckBox(GetControl('QBS_SESSIONINIT')).caption := TraduireMemoire('Session initialisée ')+BPLibelleMaille(getfield('QBS_BPINITIALISE'));
      //Ajout pour cocher la case en CWAS
      if getfield('QBS_BPINITIALISE') <> '' then THCheckBox(GetControl('QBS_SESSIONINIT')).Checked := true;

      //menu popdelai
      SetControlVisible('Semaine', DonneParamB(pb_BPListeSem));
      SetControlVisible('Quinzaine', DonneParamB(pb_BPListeQuniz));
      SetControlVisible('Mois', DonneParamB(pb_BPListeMois));
      SetControlVisible('Mois445', DonneParamB(pb_BPListeMois445));
      SetControlVisible('Trimestre', DonneParamB(pb_BPListeTrim));
      SetControlVisible('Quadrimestre', DonneParamB(pb_BPListeQuadri));

      SetControlEnabled('binsert',FALSE);
      SetControlEnabled('bdelete',FALSE);
      ds.Edit;

    end;
  end ;
end;

procedure TOM_QBPSESSIONBP.OnArgument ( StArgument : String ) ;
var
  i:integer;
begin
  Inherited ;
  if Ecran.Name='QUFPBPSESSIONENT' then
  begin
    ChampValAffOK:='';
    Loaded := false;
    OkToClose := true;
    //CEGID-CCMX le 30/10/2006 DEBUT
    sArg := GetArgumentvalue(StArgument, 'ACTION');
    sOrigine := GetArgumentvalue(StArgument, 'ORIGINE');
    if sArg = 'MODIFICATION' then FTypeAction := taModif;
    if sArg = 'CREATION' then FTypeAction := taCreat;
    if sArg = 'CONSULTATION' then FTypeAction := taConsult;
    //CEGID-CCMX le 30/10/2006 FIN
    //CEGID-CCMX Le 26/01/2007 FQ N° 14682 DEBUT
    if (sOrigine = 'TYPEDA') then
    begin
      setControlEnabled( 'QBS_VUEARBRE', False);
      setControlEnabled( 'COMBOBOXAXE1', False);
      //CEGID-CCMX le 12/04/2007 FQ N° 14840
      if sArg = 'CREATION' then
        sArguments := StArgument;
      //CEGID-CCMX FQ N° 14840 FIN
    end;
    //CEGID-CCMX Le 26/01/2007 FQ N° 14682
    TToolbarButton97(GetControl('BTNINITG')).OnClick := BTNINITG_OnClick;
    TToolbarButton97(GetControl('BTNDUPLIQUER')).OnClick := BTNDUPLIQUER_OnClick;
    ThRadioGroup(GetControl('RADIOGRP')).OnClick := RADIOGRP_OnClick;
    ThEdit(GetControl('JOURD1')).OnChange := JOURD1_OnChange;
    THPopupMenu(GetControl('POPLOI')).OnPopup := POPLOI_OnPopup;
    THPopupMenu(GetControl('POPSESSION')).OnPopup := POPSESSION_OnPopup;
    THPopupMenu(GetControl('POPINITDELAI')).OnPopup := POPINITDELAI_OnPopup;
    for i := 1 to 10 do
    begin
      THValComboBox(GetControl('COMBOBOXAXE'+IntToStr(i))).OnChange := AxeAnalyse_OnChange;
      THValComboBox(GetControl('COMBOBOXAXE'+IntToStr(i))).OnEnter := AxeAnalyse_OnEnter;
      THDBValComboBox(GetControl('QBS_CODEAXES'+IntToStr(i))).OnChange := AxeAnalyse_OnLoad;
    end;
    for i := 1 to 4 do
    begin
      THDBValComboBox(GetControl('QBS_CODEAXE'+IntToStr(i))).OnEnter := AxeRestriction_OnEnter;
      THDBValComboBox(GetControl('QBS_CODEAXE'+IntToStr(i))).OnChange := AxeRestriction_OnChange;
    end;
    for i := 1 to 4 do
    begin
      ThDBEdit(GetControl('QBS_VALEURAXE'+IntToStr(i))).OnElipsisClick := ValeurAxeRestriction_OnClick;
      ThDBEdit(GetControl('QBS_VALEURAXE'+IntToStr(i))).OnDblClick := ValeurAxeRestriction_OnClick;
    end;

    THPageControl2(GetControl('PAGES' )).OnChange:=PageChange;

    //Onglet "Valeurs Affichées"
    for i := 1 to 7 do
    begin
      ThDBValComboBox(GetControl('QBS_VALAFFH'+IntToStr(i))).OnChange := ValAff_OnChange;
      ThDBValComboBox(GetControl('QBS_VALAFFR'+IntToStr(i))).OnChange := ValAff_OnChange;
      ThDBValComboBox(GetControl('QBS_VALAFFH'+IntToStr(i))).OnEnter := ValAff_OnEnter;
      ThDBValComboBox(GetControl('QBS_VALAFFR'+IntToStr(i))).OnEnter := ValAff_OnEnter;

      ThValComboBox(GetControl('QBS_VALAFFH'+IntToStr(i)+'A_')).OnChange := ValRub_OnChange;
      ThValComboBox(GetControl('QBS_VALAFFH'+IntToStr(i)+'A_')).OnEnter := ValRub_OnEnter;
      ThValComboBox(GetControl('QBS_VALAFFR'+IntToStr(i)+'A_')).OnChange := ValRub_OnChange;
      ThValComboBox(GetControl('QBS_VALAFFR'+IntToStr(i)+'A_')).OnEnter := ValRub_OnEnter;

    end;

    Case ContextBP of
      0,1 : begin //Mode-GC
            SetControlProperty('QBS_NATURECMD','DATATYPE','GCNATUREPIECEG');
            SetControlProperty('QBS_VUEARBRE','DATATYPE','QUTBPVUEARBRE001');
            SetControlVisible('QBS_NATURE2',False);
            SetControlVisible('TQBS_NATURE2',False);
            THTabSheet(GetControl('VALAFF')).TabVisible := false;
          end;
      2 : begin //Compta
            SetControlProperty('QBS_NATURECMD','DATATYPE','TTNATUREPIECE');
            SetControlProperty('QBS_VUEARBRE','DATATYPE','QUTBPANALYTIQVALAFF');
            SetControlVisible('QBS_NATURE2',True);
            SetControlVisible('TQBS_NATURE2',True);
            THTabSheet(GetControl('VALAFF')).TabVisible := false;
          end;
      3 : begin //Paie
            SetControlVisible('QBS_NATURECMD',False);
            SetControlVisible('TQBS_NATURECMD',False);
            SetControlVisible('QBS_NATURE2',False);
            SetControlVisible('TQBS_NATURE2',False);
            SetControlVisible('QBS_VUEARBRE',False);
            SetControlVisible('TQBS_VUEARBRE',False);
            SetControlVisible('QBS_VUEARBRE',False);
            THTabSheet(GetControl('VALAFF')).TabVisible := true;
          end;
    end;

    SetControlText('EDTDUP', TrouveArgument(StArgument,'DUPLICATION',''));
    SetControlText('EDTDUPLOI', TrouveArgument(StArgument,'DUPLLOI',''));
    if ContextBP in [0,2,3]
    then SetControlVisible('BOX12', true)
    else SetControlVisible('BOX12', false);

    if v_pgi.RequeteMultiDossier.count = 0 then  THMenuItem(GetControl('MultiSocS')).Visible:=false;
  end ;
end;

procedure TOM_QBPSESSIONBP.BTNINITG_OnClick(Sender: TObject);
var InitValue : Boolean;
Caption : hString;
begin
  if THRadioGroup(GetControl('radiogrp')).itemindex = 0 then
  begin
    InitValue := TCheckbox(GetControl('QBS_SESSIONINIT')).Checked;
    Caption := TCheckbox(GetControl('QBS_SESSIONINIT')).Caption;
    SetControlChecked('QBS_SESSIONINIT',True);
    SetControlCaption('QBS_SESSIONINIT',TraduireMemoire('Session initialisée ')+FuncSynBPLibelleMaille(['4'],0));
    TFFiche(Ecran).Bouge(nbPost) ;
    if not FuncSynBPLanceCalculInitObjectif([GetControlText('QBS_CODESESSION'),'4','','',GetControlText('JOURD1'),GetControlText('JOURD1_'),GetControlText('JOURD2'),GetControlText('JOURD2_')],0)
    then
    begin
      SetControlChecked('QBS_SESSIONINIT',InitValue);
      SetControlCaption('QBS_SESSIONINIT',Caption);
    end;
  end
  else
  begin
    InitValue := TCheckbox(GetControl('QBS_SESSIONINIT')).Checked;
    Caption := TCheckbox(GetControl('QBS_SESSIONINIT')).Caption;
    SetControlChecked('QBS_SESSIONINIT',True);
    SetControlCaption('QBS_SESSIONINIT',TraduireMemoire('Session initialisée ')+FuncSynBPLibelleMaille(['0'],0));
    TFFiche(Ecran).Bouge(nbPost) ;
    if not FuncSynBPLanceCalculInitObjectif([GetControlText('QBS_CODESESSION'),'0','','',GetControlText('JOURD1'),GetControlText('JOURD1_'),GetControlText('JOURD2'),GetControlText('JOURD2_')],0)
    then
    begin
      SetControlChecked('QBS_SESSIONINIT',InitValue);
      SetControlCaption('QBS_SESSIONINIT',Caption);
    end;
  end;
  TraitementRadioGroup;
end;

procedure TOM_QBPSESSIONBP.POPLOI_OnPopup(Sender: TObject);
begin
  THMenuItem(GetControl('ajoutloi')).OnClick := ajoutloi_OnClick;
  THMenuItem(GetControl('calculloi')).OnClick := calculloi_OnClick;
end;

procedure TOM_QBPSESSIONBP.ajoutloi_OnClick(Sender: TObject);
begin
  TFFiche(Ecran).Bouge(nbPost) ;
  //CEGID-CCMX le 02/11/2006 : Gestion de l'origine TypeDA
  if (sOrigine = 'TYPEDA') then
      AglLanceFiche('Q', 'QUFMBPLOIPGI', '', '','SESSION='+GetControlText('QBS_CODESESSION')+
                                           ';DATEDEB='+GetControlText('JOURD1')+
                                           ';DATEFIN='+GetControlText('JOURD1_')+
                                           ';DATEDEBREF='+GetControlText('JOURD2')+
                                           ';DATEFINREF='+GetControlText('JOURD2_')+
                                           ';ORIGINE='+sOrigine)
  else
    AglLanceFiche('Q', 'QUFMBPLOIPGI', '', '','SESSION='+GetControlText('QBS_CODESESSION')+
                                           ';DATEDEB='+GetControlText('JOURD1')+
                                           ';DATEFIN='+GetControlText('JOURD1_')+
                                           ';DATEDEBREF='+GetControlText('JOURD2')+
                                           ';DATEFINREF='+GetControlText('JOURD2_'));
end;

procedure TOM_QBPSESSIONBP.calculloi_OnClick(Sender: TObject);
begin
  TFFiche(Ecran).Bouge(nbPost) ;
  ProcSynBPLanceCalculEclatementLoi([GetControlText('QBS_CODESESSION'),'DELAI',GetControlText('JOURD1'),GetControlText('JOURD1_'),GetControlText('JOURD2'),GetControlText('JOURD2_')],0);
end;

procedure TOM_QBPSESSIONBP.POPSESSION_OnPopup(Sender: TObject);
begin
  THMenuItem(GetControl('ControleS')).OnClick := ControleS_OnClick;
  THMenuItem(GetControl('Sdetaillee')).OnClick := Sdetaillee_OnClick;
  THMenuItem(GetControl('MultiSocS')).OnClick := MultiSocS_OnClick;
  THMenuItem(GetControl('ValideS')).OnClick := ValideS_OnClick;
  THMenuItem(GetControl('InvalideS')).OnClick := InvalideS_OnClick;
end;

procedure TOM_QBPSESSIONBP.Sdetaillee_OnClick(Sender: TObject);
var code : hString;
begin
  TFFiche(Ecran).Bouge(nbPost) ;
  if FuncSynBPSessionValide([GetControlText('QBS_CODESESSION')],0)
  then code:='ACTION=CONSULTATION;SESSION='+GetControlText('QBS_CODESESSION')+';CONSULT=CONSULTATION'
  else code:='ACTION=MODIFICATION;SESSION='+GetControlText('QBS_CODESESSION')+';CONSULT=NON';
  if FuncSynBPVerifTree([GetControlText('QBS_CODESESSION')],0)
  then AGLLanceFiche('Q','QUFSLBParbre','SESSION='+GetControlText('QBS_CODESESSION')+';NUMNOEUDPERE=0','',code)
  else HShowMessage('1;Erreur;La session ne contient aucune donnée.;W;O;O;O;','','')
end;

procedure TOM_QBPSESSIONBP.ControleS_OnClick(Sender: TObject);
begin
  TFFiche(Ecran).Bouge(nbPost) ;
  if FuncSynBPVerifTree([GetControlText('QBS_CODESESSION')],0)
  then   AglLanceFiche('Q','QUFSBPARBREPREV','','','SESSION='+GetControlText('QBS_CODESESSION')+';VUEPGIDELAI=;VUECOEFF=NON')
  else HShowMessage('1;Erreur;La session ne contient aucune donnée.;W;O;O;O;','','')
end;

procedure TOM_QBPSESSIONBP.POPINITDELAI_OnPopup(Sender: TObject);
begin
  THMenuItem(GetControl('Mois445')).OnClick := Mois445_OnClick;
  THMenuItem(GetControl('Semaine')).OnClick := Semaine_OnClick;
  THMenuItem(GetControl('Quinzaine')).OnClick := Quinzaine_OnClick;
  THMenuItem(GetControl('Mois')).OnClick := Mois_OnClick;
  THMenuItem(GetControl('Trimestre')).OnClick := Trimestre_OnClick;
  THMenuItem(GetControl('Quadrimestre')).OnClick := Quadrimestre_OnClick;
end;

procedure TOM_QBPSESSIONBP.Mois445_OnClick(Sender: TObject) ;
var InitValue : Boolean;
Caption : hString;
begin
  InitValue := TCheckbox(GetControl('QBS_SESSIONINIT')).Checked;
  Caption := TCheckbox(GetControl('QBS_SESSIONINIT')).Caption;
  SetControlChecked('QBS_SESSIONINIT',True);
  SetControlCaption('QBS_SESSIONINIT',TraduireMemoire('Session initialisée ')+FuncSynBPLibelleMaille(['5'],0));
  TFFiche(Ecran).Bouge(nbPost) ;
  if not FuncSynBPLanceCalculInitObjectif([GetControlText('QBS_CODESESSION'),'5','','',GetControlText('JOURD1'),GetControlText('JOURD1_'),GetControlText('JOURD2'),GetControlText('JOURD2_')],0)
  then
  begin
    SetControlChecked('QBS_SESSIONINIT',InitValue);
    SetControlCaption('QBS_SESSIONINIT',Caption);
  end;
end;

procedure TOM_QBPSESSIONBP.Semaine_OnClick(Sender: TObject);
var InitValue : Boolean;
Caption : hString;
begin
  InitValue := TCheckbox(GetControl('QBS_SESSIONINIT')).Checked;
  Caption := TCheckbox(GetControl('QBS_SESSIONINIT')).Caption;
  SetControlChecked('QBS_SESSIONINIT',True);
  SetControlCaption('QBS_SESSIONINIT',TraduireMemoire('Session initialisée ')+FuncSynBPLibelleMaille(['2'],0));
  TFFiche(Ecran).Bouge(nbPost) ;
  if not FuncSynBPLanceCalculInitObjectif([GetControlText('QBS_CODESESSION'),'2','','',GetControlText('JOURD1'),GetControlText('JOURD1_'),GetControlText('JOURD2'),GetControlText('JOURD2_')],0)
  then
  begin
    SetControlChecked('QBS_SESSIONINIT',InitValue);
    SetControlCaption('QBS_SESSIONINIT',Caption);
  end;
end;

procedure TOM_QBPSESSIONBP.Quinzaine_OnClick(Sender: TObject);
var InitValue : Boolean;
Caption : hString;
begin
  InitValue := TCheckbox(GetControl('QBS_SESSIONINIT')).Checked;
  Caption := TCheckbox(GetControl('QBS_SESSIONINIT')).Caption;
  SetControlChecked('QBS_SESSIONINIT',True);
  SetControlCaption('QBS_SESSIONINIT',TraduireMemoire('Session initialisée ')+FuncSynBPLibelleMaille(['3'],0));
  TFFiche(Ecran).Bouge(nbPost);
  if not FuncSynBPLanceCalculInitObjectif([GetControlText('QBS_CODESESSION'),'3','','',GetControlText('JOURD1'),GetControlText('JOURD1_'),GetControlText('JOURD2'),GetControlText('JOURD2_')],0)
  then
  begin
    SetControlChecked('QBS_SESSIONINIT',InitValue);
    SetControlCaption('QBS_SESSIONINIT',Caption);
  end;
end;

procedure TOM_QBPSESSIONBP.Mois_OnClick(Sender: TObject);
var InitValue : Boolean;
Caption : hString;
begin
  InitValue := TCheckbox(GetControl('QBS_SESSIONINIT')).Checked;
  Caption := TCheckbox(GetControl('QBS_SESSIONINIT')).Caption;
  SetControlChecked('QBS_SESSIONINIT',True);
  SetControlCaption('QBS_SESSIONINIT',TraduireMemoire('Session initialisée ')+FuncSynBPLibelleMaille(['4'],0));
  TFFiche(Ecran).Bouge(nbPost) ;
  if not FuncSynBPLanceCalculInitObjectif([GetControlText('QBS_CODESESSION'),'4','','',GetControlText('JOURD1'),GetControlText('JOURD1_'),GetControlText('JOURD2'),GetControlText('JOURD2_')],0)
  then
  begin
    SetControlChecked('QBS_SESSIONINIT',InitValue);
    SetControlCaption('QBS_SESSIONINIT',Caption);
  end;
end;

procedure TOM_QBPSESSIONBP.Trimestre_OnClick(Sender: TObject);
var InitValue : Boolean;
Caption : hString;
begin
  InitValue := TCheckbox(GetControl('QBS_SESSIONINIT')).Checked;
  Caption := TCheckbox(GetControl('QBS_SESSIONINIT')).Caption;
  SetControlChecked('QBS_SESSIONINIT',True);
  SetControlCaption('QBS_SESSIONINIT',TraduireMemoire('Session initialisée ')+FuncSynBPLibelleMaille(['6'],0));
  TFFiche(Ecran).Bouge(nbPost) ;
  if not FuncSynBPLanceCalculInitObjectif([GetControlText('QBS_CODESESSION'),'6','','',GetControlText('JOURD1'),GetControlText('JOURD1_'),GetControlText('JOURD2'),GetControlText('JOURD2_')],0)
  then
  begin
    SetControlChecked('QBS_SESSIONINIT',InitValue);
    SetControlCaption('QBS_SESSIONINIT',Caption);
  end;
end;

procedure TOM_QBPSESSIONBP.Quadrimestre_OnClick(Sender: TObject);
var InitValue : Boolean;
Caption : hString;
begin
  InitValue := TCheckbox(GetControl('QBS_SESSIONINIT')).Checked;
  Caption := TCheckbox(GetControl('QBS_SESSIONINIT')).Caption;
  SetControlChecked('QBS_SESSIONINIT',True);
  SetControlCaption('QBS_SESSIONINIT',TraduireMemoire('Session initialisée ')+FuncSynBPLibelleMaille(['7'],0));
  TFFiche(Ecran).Bouge(nbPost) ;
  if not FuncSynBPLanceCalculInitObjectif([GetControlText('QBS_CODESESSION'),'7','','',GetControlText('JOURD1'),GetControlText('JOURD1_'),GetControlText('JOURD2'),GetControlText('JOURD2_')],0)
  then
  begin
    SetControlChecked('QBS_SESSIONINIT',InitValue);
    SetControlCaption('QBS_SESSIONINIT',Caption);
  end;
end;

//pour les 10 axes analyses
procedure TOM_QBPSESSIONBP.AxeAnalyse_OnChange(Sender: TObject);
var NameAxe:string;
begin
  NameAxe := TControl( Sender ).name;
  if NameAxe <> 'COMBOBOXAXE10' then
  begin
    if GetField('QBS_CODEAXES'+AnsiRightStr(NameAxe,1)) <> GetControlText(NameAxe) then
       SetField('QBS_CODEAXES'+AnsiRightStr(NameAxe,1),GetControlText(NameAxe))
  end
  else
  begin
    if GetField('QBS_CODEAXES10') <> GetControlText(NameAxe) then SetField('QBS_CODEAXES10',GetControlText(NameAxe))
  end;
end;

procedure TOM_QBPSESSIONBP.AxeAnalyse_OnLoad(Sender: TObject);
var NameAxe:string;
begin
  //FQ MODE 15088
  if Loaded = false then exit;
  NameAxe := TControl( Sender ).name;
  if NameAxe <> 'QBS_CODEAXES10' then SetControlText('COMBOBOXAXE'+AnsiRightStr(NameAxe,1),GetControlText('QBS_CODEAXES'+AnsiRightStr(NameAxe,1)))
  else SetControlText('COMBOBOXAXE10',GetControlText('QBS_CODEAXES10'))
end;


procedure TOM_QBPSESSIONBP.AxeAnalyse_OnEnter(Sender: TObject);
var NameAxe:string;
begin
  NameAxe := TControl( Sender ).name;
  SetControlProperty(NameAxe,'Plus',GetControlText('TabPlus'));
end;

procedure TOM_QBPSESSIONBP.RADIOGRP_OnClick(Sender: TObject);
begin
  if FTypeAction <> taCreat then TraitementRadioGroup;
end;

procedure TOM_QBPSESSIONBP.JOURD1_OnChange(Sender: TObject);
var Date : TDateTime;
TestDateOk : hString;
begin
  { EVI / EConvertError pendant la saisie }
  TestDateOk := GetControlText('jourd1');
  if Pos(' ',TestDateOk) = 0 then
  begin
    if Loaded = true then
    begin
      Date := STRTODATE(GetControlText('jourd1'))+(STRTODATE(GetControlText('jourd2_'))-STRTODATE(GetControlText('jourd2')));
      SetControlText('jourd1_',DateToStr(Date));
    end;
  end;
end;

//pour les 4 axes de restrictions
procedure TOM_QBPSESSIONBP.AxeRestriction_OnEnter(Sender: TObject);
var NameAxe:string;
begin
  NameAxe := TControl( Sender ).name;
  SetControlProperty(NameAxe,'Plus',GetControlText('TabPlus'));
end;

procedure TOM_QBPSESSIONBP.AxeRestriction_OnChange(Sender: TObject);
var NameAxe:string;
begin
  NameAxe := TControl( Sender ).name;
  SetControlText('QBS_VALEURAXE'+AnsiRightStr(NameAxe,1),'');
  SetControlProperty('QBS_VALEURAXE'+AnsiRightStr(NameAxe,1),'DataType',FuncSynBPTabletteValeurAxeC([GetControlText(NameAxe)],0));
  SetControlProperty('QBS_VALEURAXE'+AnsiRightStr(NameAxe,1),'Plus',FuncSynBPPlusValeurAxe([GetControlText(NameAxe)],0));
end;

procedure TOM_QBPSESSIONBP.ValideS_OnClick(Sender: TObject);
begin
  TFFiche(Ecran).Bouge(nbPost) ;
  if (not FuncSynBPSessionValide([GetField('QBS_CODESESSION')],0))
  then ProcSynBPLanceValidation([GetField('QBS_CODESESSION')],0);
  SetControlText('edtferme','OUI');
  Ecran.Close;
end;

procedure TOM_QBPSESSIONBP.InvalideS_OnClick(Sender: TObject);
begin
  TFFiche(Ecran).Bouge(nbPost) ;
  if (FuncSynBPSessionValide([GetField('QBS_CODESESSION')],0))
  then ProcSynBPLanceDeValidation([GetField('QBS_CODESESSION')],0);
  SetControlText('edtferme','OUI');
  Ecran.Close;
end;

procedure TOM_QBPSESSIONBP.BTNDUPLIQUER_OnClick(Sender: TObject);
var retour, dup, code, session, codesession : hString;
Q : TQuery;
i : integer;
begin
  SetControlText('okDupTotale','NON');
  SetControlText('NivDupPartielle','');
  retour:=AglLanceFiche('Q','QUFVBPDUPLIQUER','','',GetField('QBS_CODESESSION'));
  dup:='';

  //Vérification de l'existence de la session dans la base
  codesession := TrouveArgument(retour,'SESSION','');
  if retour<>'' then
  begin
    if not ExisteSQL('SELECT QBS_CODESESSION FROM QBPSESSIONBP WHERE QBS_CODESESSION="'+codesession+'"')
    then
    begin
      if codesession = '' then HShowMessage('1;Erreur;Le code session est vide.;W;O;O;O;','','')
      else HShowMessage('1;Erreur;'+codesession+' n''est pas une session valide.;W;O;O;O;','','');
      //Réinitialisation de la valeur retour
      retour := '';
    end;
  end;

  if retour<>'' then
  begin
    dup:=TrouveArgument(retour,'DUP','');
    SetControlText('SessionDup',TrouveArgument(retour,'SESSION',''));
    SetControlText('edtokloi','7');
    if (dup<>'1') and (dup<>'2') then
    if (FuncSynBPBtnInitDelaiObj([GetControlText('SessionDup')],0))=false
    then SetControlText('edtokloi', inttostr(HShowMessage('0;Session Objectif;Voulez-vous aussi dupliquer les lois ?;Q;YN;N;N;','','')) );

    code:='';
    if (dup='1') or (dup='2') then code:=',QBS_SESSIONINIT,QBS_SESSIONECLAT,QBS_INITDELAI,QBS_SESSIONVALIDE ';

    Q:= openSql('SELECT QBS_LIBSESSION,QBS_DATEDEBC,QBS_DATEFINC,QBS_DATEDEBREF,QBS_DATEFINREF,QBS_DATEOUVERT,'+
                'QBS_CODEAXE1,QBS_VALEURAXE1,QBS_CODEAXE2,QBS_VALEURAXE2,QBS_CODEAXE3,QBS_VALEURAXE3,'+
                'QBS_CODEAXE4,QBS_VALEURAXE4,QBS_SAISONCMDC,QBS_SAISONCMDREF,QBS_NATURECMD,QBS_NATURE2,QBS_MODEGESTPREV,'+
                'QBS_EXTRAPOLABLE,QBS_SESSIONECLAT,QBS_TYPENATURE,QBS_CODEAXES1,QBS_CODEAXES2,QBS_CODEAXES3,'+
                'QBS_CODEAXES4,QBS_CODEAXES5,QBS_CODEAXES6,QBS_CODEAXES7,QBS_CODEAXES8,QBS_CODEAXES9,'+
                'QBS_CODEAXES10,QBS_NATURE2,QBS_NBVALAFF,QBS_VALAFFH1,QBS_VALAFFH1A,QBS_VALAFFH1B,'+
                'QBS_VALAFFR1,QBS_VALAFFR1A,QBS_VALAFFR1B,QBS_VALAFFH2,QBS_VALAFFH2A,QBS_VALAFFH2B,'+
                'QBS_VALAFFR2,QBS_VALAFFR2A,QBS_VALAFFR2B,QBS_VALAFFH3,QBS_VALAFFH3A,QBS_VALAFFH3B,'+
                'QBS_VALAFFR3,QBS_VALAFFR3A,QBS_VALAFFR3B,QBS_VALAFFH4,QBS_VALAFFH4A,QBS_VALAFFH4B,'+
                'QBS_VALAFFR4,QBS_VALAFFR4A,QBS_VALAFFR4B,QBS_VALAFFH5,QBS_VALAFFH5A,QBS_VALAFFH5B,'+
                'QBS_VALAFFR5,QBS_VALAFFR5A,QBS_VALAFFR5B,QBS_VALAFFH6,QBS_VALAFFH6A,QBS_VALAFFH6B,'+
                'QBS_VALAFFR6,QBS_VALAFFR6A,QBS_VALAFFR6B,QBS_VALAFFH7,QBS_VALAFFH7A,QBS_VALAFFH7B,'+
                'QBS_VALAFFR7,QBS_VALAFFR7A,QBS_VALAFFR7B,QBS_VALAFFLIB1,QBS_VALAFFLIB2,QBS_VALAFFLIB3,'+
                'QBS_VALAFFLIB4,QBS_VALAFFLIB5,QBS_VALAFFLIB6,QBS_VALAFFLIB7,QBS_CODESTRUCT,QBS_VUEARBRE,'+
                'QBS_METHODE,QBS_BPINITIALISE '+code+
                ' FROM QBPSESSIONBP WHERE QBS_CODESESSION="'+GetControlText('SessionDup')+'"',true);
    session:=IncrementeCodeSession(GetControlText('SessionDup'));
    SetField('QBS_CODESESSION',session);
    SetField('QBS_LIBSESSION',Q.Fields[0].AsString);
    SetField('QBS_DATEDEBC',Q.Fields[1].AsDateTime);
    SetField('QBS_DATEFINC',Q.Fields[2].AsDateTime);
    SetField('QBS_DATEDEBREF',Q.Fields[3].AsDateTime);
    SetField('QBS_DATEFINREF',Q.Fields[4].AsDateTime);
    SetField('QBS_DATEOUVERT',Q.Fields[5].AsDateTime);
    SetField('QBS_CODEAXE1',Q.Fields[6].AsString);
    SetField('QBS_VALEURAXE1',Q.Fields[7].AsString);
    SetField('QBS_CODEAXE2',Q.Fields[8].AsString);
    SetField('QBS_VALEURAXE2',Q.Fields[9].AsString);
    SetField('QBS_CODEAXE3',Q.Fields[10].AsString);
    SetField('QBS_VALEURAXE3',Q.Fields[11].AsString);
    SetField('QBS_CODEAXE4',Q.Fields[12].AsString);
    SetField('QBS_VALEURAXE4',Q.Fields[13].AsString);
    SetField('QBS_SAISONCMDC',Q.Fields[14].AsString);
    SetField('QBS_SAISONCMDREF',Q.Fields[15].AsString);
    SetField('QBS_NATURECMD',Q.Fields[16].AsString);
    SetField('QBS_NATURE2',Q.Fields[17].AsString);
    SetField('QBS_MODEGESTPREV',Q.Fields[18].AsString);
    SetField('QBS_EXTRAPOLABLE',Q.Fields[19].AsString);
    SetField('QBS_SESSIONECLAT','-');
    SetField('QBS_TYPENATURE',Q.Fields[21].AsString);
    SetField('QBS_CODEAXES1',Q.Fields[22].AsString);
    SetField('QBS_CODEAXES2',Q.Fields[23].AsString);
    SetField('QBS_CODEAXES3',Q.Fields[24].AsString);
    SetField('QBS_CODEAXES4',Q.Fields[25].AsString);
    SetField('QBS_CODEAXES5',Q.Fields[26].AsString);
    SetField('QBS_CODEAXES6',Q.Fields[27].AsString);
    SetField('QBS_CODEAXES7',Q.Fields[28].AsString);
    SetField('QBS_CODEAXES8',Q.Fields[29].AsString);
    SetField('QBS_CODEAXES9',Q.Fields[30].AsString);
    SetField('QBS_CODEAXES10',Q.Fields[31].AsString);
    SetField('QBS_NATURE2',Q.Fields[32].AsString);
    SetField('QBS_NBVALAFF',Q.Fields[33].AsInteger);
    SetField('QBS_VALAFFH1',Q.Fields[34].AsString);
    SetField('QBS_VALAFFH1A',Q.Fields[35].AsString);
    SetField('QBS_VALAFFH1B',Q.Fields[36].AsString);
    SetField('QBS_VALAFFR1',Q.Fields[37].AsString);
    SetField('QBS_VALAFFR1A',Q.Fields[38].AsString);
    SetField('QBS_VALAFFR1B',Q.Fields[39].AsString);
    SetField('QBS_VALAFFH2',Q.Fields[40].AsString);
    SetField('QBS_VALAFFH2A',Q.Fields[41].AsString);
    SetField('QBS_VALAFFH2B',Q.Fields[42].AsString);
    SetField('QBS_VALAFFR2',Q.Fields[43].AsString);
    SetField('QBS_VALAFFR2A',Q.Fields[44].AsString);
    SetField('QBS_VALAFFR2B',Q.Fields[45].AsString);
    SetField('QBS_VALAFFH3',Q.Fields[46].AsString);
    SetField('QBS_VALAFFH3A',Q.Fields[47].AsString);
    SetField('QBS_VALAFFH3B',Q.Fields[48].AsString);
    SetField('QBS_VALAFFR3',Q.Fields[49].AsString);
    SetField('QBS_VALAFFR3A',Q.Fields[50].AsString);
    SetField('QBS_VALAFFR3B',Q.Fields[51].AsString);
    SetField('QBS_VALAFFH4',Q.Fields[52].AsString);
    SetField('QBS_VALAFFH4A',Q.Fields[53].AsString);
    SetField('QBS_VALAFFH4B',Q.Fields[54].AsString);
    SetField('QBS_VALAFFR4',Q.Fields[55].AsString);
    SetField('QBS_VALAFFR4A',Q.Fields[56].AsString);
    SetField('QBS_VALAFFR4B',Q.Fields[57].AsString);
    SetField('QBS_VALAFFH5',Q.Fields[58].AsString);
    SetField('QBS_VALAFFH5A',Q.Fields[59].AsString);
    SetField('QBS_VALAFFH5B',Q.Fields[60].AsString);
    SetField('QBS_VALAFFR5',Q.Fields[61].AsString);
    SetField('QBS_VALAFFR5A',Q.Fields[62].AsString);
    SetField('QBS_VALAFFR5B',Q.Fields[63].AsString);
    SetField('QBS_VALAFFH6',Q.Fields[64].AsString);
    SetField('QBS_VALAFFH6A',Q.Fields[65].AsString);
    SetField('QBS_VALAFFH6B',Q.Fields[66].AsString);
    SetField('QBS_VALAFFR6',Q.Fields[67].AsString);
    SetField('QBS_VALAFFR6A',Q.Fields[68].AsString);
    SetField('QBS_VALAFFR6B',Q.Fields[69].AsString);
    SetField('QBS_VALAFFH7',Q.Fields[70].AsString);
    SetField('QBS_VALAFFH7A',Q.Fields[71].AsString);
    SetField('QBS_VALAFFH7B',Q.Fields[72].AsString);
    SetField('QBS_VALAFFR7',Q.Fields[73].AsString);
    SetField('QBS_VALAFFR7A',Q.Fields[74].AsString);
    SetField('QBS_VALAFFR7B',Q.Fields[75].AsString);
    SetField('QBS_VALAFFLIB1',Q.Fields[76].AsString);
    SetField('QBS_VALAFFLIB2',Q.Fields[77].AsString);
    SetField('QBS_VALAFFLIB3',Q.Fields[78].AsString);
    SetField('QBS_VALAFFLIB4',Q.Fields[79].AsString);
    SetField('QBS_VALAFFLIB5',Q.Fields[80].AsString);
    SetField('QBS_VALAFFLIB6',Q.Fields[81].AsString);
    SetField('QBS_VALAFFLIB7',Q.Fields[82].AsString);
    SetField('QBS_CODESTRUCT',Q.Fields[83].AsString);
    SetField('QBS_VUEARBRE',Q.Fields[84].AsString);
    SetField('QBS_METHODE',Q.Fields[85].AsString);
    if (dup='1') or (dup='2') then
    begin
      SetField('QBS_BPINITIALISE',Q.Fields[86].AsString);
      SetField('QBS_SESSIONINIT',Q.Fields[87].AsString);
      SetField('QBS_SESSIONECLAT',Q.Fields[88].AsString);
      SetField('QBS_INITDELAI',Q.Fields[89].AsString);
      SetField('QBS_SESSIONVALIDE',Q.Fields[90].AsString);
    end;
    Ferme(Q);

    SetControlText('Jourd1', DateToStr(strToDateTime0(GetControlText('QBS_DATEDEBC'))));
    SetControlText('Jourd1_', DateToStr(strToDateTime0(GetControlText('QBS_DATEFINC'))));
    SetControlText('Jourd2', DateToStr(strToDateTime0(GetControlText('QBS_DATEDEBREF'))));
    SetControlText('Jourd2_', DateToStr(strToDateTime0(GetControlText('QBS_DATEFINREF'))));

    if GetField('QBS_METHODE')='1'
    then SetControlProperty('radiogrp','ITEMINDEX',0)
    else SetControlProperty('radiogrp','ITEMINDEX',1);
  end;

  //totale
  if dup='1' then
  begin
    SetControlText('okDupTotale', 'OUI');
    SetControlEnabled('QBS_CODEAXE1', false);
    SetControlEnabled('QBS_VALEURAXE1', false);
    SetControlEnabled('QBS_CODEAXE2', false);
    SetControlEnabled('QBS_VALEURAXE2', false);
    SetControlEnabled('QBS_CODEAXE3', false);
    SetControlEnabled('QBS_VALEURAXE3', false);
    SetControlEnabled('QBS_CODEAXE4', false);
    SetControlEnabled('QBS_VALEURAXE4', false);
    SetControlEnabled('QBS_NATURECMD', false);
    SetControlEnabled('QBS_NATURE2', false);
    SetControlEnabled('QBS_CODEAXES1', false);
    SetControlEnabled('QBS_CODEAXES2', false);
    SetControlEnabled('QBS_CODEAXES3', false);
    SetControlEnabled('QBS_CODEAXES4', false);
    SetControlEnabled('QBS_CODEAXES5', false);
    SetControlEnabled('QBS_CODEAXES6', false);
    SetControlEnabled('QBS_CODEAXES7', false);
    SetControlEnabled('QBS_CODEAXES8', false);
    SetControlEnabled('QBS_CODEAXES9', false);
    SetControlEnabled('QBS_CODEAXES10', false);

    SetControlEnabled('QBS_VUEARBRE', false);

    SetControlText('edtokloi', '6');

    SetControlEnabled('Jourd1', false);
    SetControlEnabled('Jourd1_', false);
    SetControlEnabled('Jourd2', false);
    SetControlEnabled('Jourd2_', false);
    SetControlEnabled('radiogrp', false);
    { EVI / Bloque la possibilité de changer la valeur à afficher une fois l'initialisation faite }
    SetControlEnabled('QBS_VUEARBRE', FALSE);
  end;

  //partielle
  if dup='2' then
  begin
    SetControlText('NivDupPartielle', TrouveArgument(retour,'NIV',''));
    for i := 9 downto StrToInt(GetControlText('NivDupPartielle'))+1 do SetField('QBS_CODEAXES'+IntToStr(i+1),'');
    SetControlEnabled('QBS_CODEAXE1', false);
    SetControlEnabled('QBS_VALEURAXE1', false);
    SetControlEnabled('QBS_CODEAXE2', false);
    SetControlEnabled('QBS_VALEURAXE2', false);
    SetControlEnabled('QBS_CODEAXE3', false);
    SetControlEnabled('QBS_VALEURAXE3', false);
    SetControlEnabled('QBS_CODEAXE4', false);
    SetControlEnabled('QBS_VALEURAXE4', false);
    SetControlEnabled('QBS_NATURECMD', false);
    SetControlEnabled('QBS_NATURE2', false);
    SetControlEnabled('QBS_CODEAXES1', false);
    SetControlEnabled('QBS_CODEAXES2', false);
    SetControlENabled('QBS_CODEAXES3', false);
    SetControlEnabled('QBS_CODEAXES4', false);
    SetControlEnabled('QBS_CODEAXES5', false);
    SetControlEnabled('QBS_CODEAXES6', false);
    SetControlEnabled('QBS_CODEAXES7', false);
    SetControlEnabled('QBS_CODEAXES8', false);
    SetControlEnabled('QBS_CODEAXES9', false);
    SetControlEnabled('QBS_CODEAXES10', false);
    SetControlEnabled('QBS_VUEARBRE', false);

    SetControlText('edtokloi','6');
    SetControlEnabled('Jourd1', false);
    SetControlEnabled('Jourd1_', false);
    SetControlEnabled('Jourd2', false);
    SetControlEnabled('Jourd2_', false);
    SetControlEnabled('radiogrp', false);
   { EVI / Bloque la possibilité de changer la valeur à afficher une fois l'initialisation faite }
    SetControlEnabled('QBS_VUEARBRE', FALSE);
  end;
end;

procedure TOM_QBPSESSIONBP.ValeurAxeRestriction_OnClick(Sender: TObject);
var NameAxe,CodeSession:string;
begin
  NameAxe := TControl( Sender ).name;
  CodeSession := GetField('QBS_CODESESSION');
  SetField(NameAxe, FuncSynBPLanceChxValeurAxe(['multi',GetControlText('QBS_CODEAXE'+AnsiRightStr(NameAxe,1)),'[@CREATION@]'],0));
end;

Function TOM_QBPSESSIONBP.FuncBPSessionMAJChamp : boolean;
var //S3,S4,S5,S6:hString;
RadioGroup : THRadioGroup;
begin
  result := true;

  SetField('QBS_DATEDEBC',FuncSynMakeDate([GetControlText('JOURD1'),''],0)) ;
  SetField('QBS_DATEDEBREF',FuncSynMakeDate([GetControlText('JOURD2'),''],0)) ;
  SetField('QBS_DATEFINC',FuncSynMakeDate([GetControlText('JOURD1_'),''],0)) ;
  SetField('QBS_DATEFINREF',FuncSynMakeDate([GetControlText('JOURD2_'),''],0)) ;

  RadioGroup := THRadioGroup(GetControl('radiogrp'));
  if RadioGroup.itemindex = 0 then SetField('QBS_METHODE','1') ;
  if RadioGroup.itemindex = 1 then SetField('QBS_METHODE','2') ;

  Case ContextBP of
    0,2,3 : SetField('QBS_CODESTRUCT','#PGIENT');
    1 : SetField('QBS_CODESTRUCT','#PGI');
  end;

  SetField('QBS_OKMODIFVUE','X');

end;

procedure TOM_QBPSESSIONBP.ValAff_OnChange(Sender: TObject);
var ValAff:string;
begin
  ValAff := TControl( Sender ).name;
  if GetControlText(ValAff) <> LastValue then
  begin
    if GetControlText(ValAff) = 'PPU' then
    begin
      SetControlVisible(ValAff+'A_',false);
      SetControlText(ValAff+'A',TraduireMemoire('<<Tous>>'));
      SetControlProperty(ValAff+'A','DATATYPE','QUTBPPGPPU');
      SetControlVisible(ValAff+'A',true);
      SetControlVisible(ValAff+'B',false);
    end;
    if GetControlText(ValAff) = 'PHB' then
    begin
      SetControlText(ValAff+'A',TraduireMemoire('<<Tous>>'));
      SetControlVisible(ValAff+'A',false);
      SetControlVisible(ValAff+'A_',true);
      THValComboBox(GetControl(ValAff+'A_')).ItemIndex := 0;
      SetControlText(ValAff+'A_',GetControlText(ValAff+'A_'));
      SetControlText(ValAff+'B',TraduireMemoire('<<Tous>>'));
      SetControlVisible(ValAff+'B',true);
    end;
    if GetControlText(ValAff) = 'PHC' then
    begin
      SetControlVisible(ValAff+'A_',false);
      SetControlText(ValAff+'A',TraduireMemoire('<<Tous>>'));
      SetControlProperty(ValAff+'A','DATATYPE','PGCUMULPAIE');
      SetControlVisible(ValAff+'A',true);
      SetControlVisible(ValAff+'B',false);
    end;
    if GetControlText(ValAff) = 'PSA' then
    begin
      SetControlVisible(ValAff+'A_',false);
      SetControlText(ValAff+'A',TraduireMemoire('<<Tous>>'));
      SetControlProperty(ValAff+'A','DATATYPE','');
      LibellesRubSalPaie(ValAff+'A');
      SetControlVisible(ValAff+'A',true);
      SetControlVisible(ValAff+'B',false);
    end;
    if GetControlText(ValAff) = 'NBS' then
    begin
      SetControlVisible(ValAff+'A_',false);
      SetControlText(ValAff+'A',TraduireMemoire('<<Tous>>'));
      SetControlVisible(ValAff+'A',false);
      SetControlVisible(ValAff+'B',false);
    end;
    if GetControlText(ValAff) = 'EFM' then
    begin
      SetControlVisible(ValAff+'A_',false);
      SetControlText(ValAff+'A',TraduireMemoire('<<Tous>>'));
      SetControlVisible(ValAff+'A',false);
      SetControlVisible(ValAff+'B',false);
    end;
    if GetControlText(ValAff) = 'ETP' then
    begin
      SetControlVisible(ValAff+'A_',false);
      SetControlText(ValAff+'A',TraduireMemoire('<<Tous>>'));
      SetControlVisible(ValAff+'A',false);
      SetControlVisible(ValAff+'B',false);
    end;
    LastValue := '';
  end;
end;

procedure TOM_QBPSESSIONBP.ValAff_OnEnter(Sender: TObject);
var ValAff:string;
begin
  ValAff := TControl( Sender ).name;
  LastValue := GetControlText(ValAff)
end;

procedure TOM_QBPSESSIONBP.PageChange(Sender: TObject);
var TabActive:hString;
i:integer;
begin
  TabActive:=THPageControl2(GetControl('PAGES' )).ActivePage.Name;
  if TabActive = 'VALAFF' then
  begin
    TabValAffSelected := true;
    for i:=1 to ThSpinEdit(GetControl('QBS_NBVALAFF')).Value do
    begin
      if GetControlText('QBS_VALAFFH'+IntToStr(i)) <> '' then
      begin
        if GetControlText('QBS_VALAFFH'+IntToStr(i)) = 'PPU' then
        begin
          SetControlProperty('QBS_VALAFFH'+IntToStr(i)+'A','DATATYPE','QUTBPPGPPU');
          SetControlVisible('QBS_VALAFFH'+IntToStr(i)+'A_',false);
          SetControlVisible('QBS_VALAFFH'+IntToStr(i)+'A',true);
        end;
        if GetControlText('QBS_VALAFFH'+IntToStr(i)) = 'PHB' then
        begin
          SetControlVisible('QBS_VALAFFH'+IntToStr(i)+'A',false);
          SetControlVisible('QBS_VALAFFH'+IntToStr(i)+'A_',true);
          SetControlVisible('QBS_VALAFFH'+IntToStr(i)+'B',true);
          if GetControlText('QBS_VALAFFH'+IntToStr(i)+'A') = 'AAA' then
          begin
            SetControlText('QBS_VALAFFH'+IntToStr(i)+'A_','AAA');
            SetControlProperty('QBS_VALAFFH'+IntToStr(i)+'B','DATATYPE','PGREMUNERATION');
          end;
          if GetControlText('QBS_VALAFFH'+IntToStr(i)+'A') = 'COT' then
          begin
            SetControlText('QBS_VALAFFH'+IntToStr(i)+'A_','COT');
            SetControlProperty('QBS_VALAFFH'+IntToStr(i)+'B','DATATYPE','PGCOTIS');
          end;
          if GetControlText('QBS_VALAFFH'+IntToStr(i)+'A') = 'BAS' then
          begin
            SetControlText('QBS_VALAFFH'+IntToStr(i)+'A_','BAS');
            SetControlProperty('QBS_VALAFFH'+IntToStr(i)+'B','DATATYPE','PGBASECOTISATION');
          end;
        end;
        if GetControlText('QBS_VALAFFH'+IntToStr(i)) = 'PHC' then
        begin
          SetControlProperty('QBS_VALAFFH'+IntToStr(i)+'A','DATATYPE','PGCUMULPAIE');
          SetControlVisible('QBS_VALAFFH'+IntToStr(i)+'A_',false);
          SetControlVisible('QBS_VALAFFH'+IntToStr(i)+'A',true);
        end;
        if GetControlText('QBS_VALAFFH'+IntToStr(i)) = 'PSA' then
        begin
          LibellesRubSalPaie('QBS_VALAFFH'+IntToStr(i)+'A');
          SetControlVisible('QBS_VALAFFH'+IntToStr(i)+'A_',false);
          SetControlVisible('QBS_VALAFFH'+IntToStr(i)+'A',true);
        end;
        if GetControlText('QBS_VALAFFH'+IntToStr(i)) = 'NBS' then
        begin
          SetControlVisible('QBS_VALAFFH'+IntToStr(i)+'A_',false);
          SetControlVisible('QBS_VALAFFH'+IntToStr(i)+'A',false);
          SetControlVisible('QBS_VALAFFH'+IntToStr(i)+'B',false);
        end;
        if GetControlText('QBS_VALAFFH'+IntToStr(i)) = 'EFM' then
        begin
          SetControlVisible('QBS_VALAFFH'+IntToStr(i)+'A_',false);
          SetControlVisible('QBS_VALAFFH'+IntToStr(i)+'A',false);
          SetControlVisible('QBS_VALAFFH'+IntToStr(i)+'B',false);
        end;
        if GetControlText('QBS_VALAFFH'+IntToStr(i)) = 'ETP' then
        begin
          SetControlVisible('QBS_VALAFFH'+IntToStr(i)+'A_',false);
          SetControlVisible('QBS_VALAFFH'+IntToStr(i)+'A',false);
          SetControlVisible('QBS_VALAFFH'+IntToStr(i)+'B',false);
        end;
      end;
      if GetControlText('QBS_VALAFFR'+IntToStr(i)) <> '' then
      begin
        if GetControlText('QBS_VALAFFR'+IntToStr(i)) = 'PPU' then
        begin
          SetControlProperty('QBS_VALAFFR'+IntToStr(i)+'A','DATATYPE','QUTBPPGPPU');
          SetControlVisible('QBS_VALAFFR'+IntToStr(i)+'A_',false);
          SetControlVisible('QBS_VALAFFR'+IntToStr(i)+'A',true);
        end;
        if GetControlText('QBS_VALAFFR'+IntToStr(i)) = 'PHB' then
        begin
          SetControlVisible('QBS_VALAFFR'+IntToStr(i)+'A',false);
          SetControlVisible('QBS_VALAFFR'+IntToStr(i)+'A_',true);
          SetControlVisible('QBS_VALAFFR'+IntToStr(i)+'B',true);
          if GetControlText('QBS_VALAFFR'+IntToStr(i)+'A') = 'AAA' then
          begin
            SetControlText('QBS_VALAFFR'+IntToStr(i)+'A_','AAA');
            SetControlProperty('QBS_VALAFFR'+IntToStr(i)+'B','DATATYPE','PGREMUNERATION');
          end;
          if GetControlText('QBS_VALAFFR'+IntToStr(i)+'A') = 'COT' then
          begin
            SetControlText('QBS_VALAFFR'+IntToStr(i)+'A_','COT');
            SetControlProperty('QBS_VALAFFR'+IntToStr(i)+'B','DATATYPE','PGCOTIS');
          end;
          if GetControlText('QBS_VALAFFR'+IntToStr(i)+'A') = 'BAS' then
          begin
            SetControlText('QBS_VALAFFR'+IntToStr(i)+'A_','BAS');
            SetControlProperty('QBS_VALAFFR'+IntToStr(i)+'B','DATATYPE','PGBASECOTISATION');
          end;
        end;
        if GetControlText('QBS_VALAFFR'+IntToStr(i)) = 'PHC' then
        begin
          SetControlProperty('QBS_VALAFFR'+IntToStr(i)+'A','DATATYPE','PGCUMULPAIE');
          SetControlVisible('QBS_VALAFFR'+IntToStr(i)+'A_',false);
          SetControlVisible('QBS_VALAFFR'+IntToStr(i)+'A',true);
        end;
        if GetControlText('QBS_VALAFFR'+IntToStr(i)) = 'NBS' then
        begin
          SetControlVisible('QBS_VALAFFR'+IntToStr(i)+'A_',false);
          SetControlVisible('QBS_VALAFFR'+IntToStr(i)+'A',false);
          SetControlVisible('QBS_VALAFFR'+IntToStr(i)+'B',false);
        end;
        if GetControlText('QBS_VALAFFR'+IntToStr(i)) = 'EFM' then
        begin
          SetControlVisible('QBS_VALAFFR'+IntToStr(i)+'A_',false);
          SetControlVisible('QBS_VALAFFR'+IntToStr(i)+'A',false);
          SetControlVisible('QBS_VALAFFR'+IntToStr(i)+'B',false);
        end;
        if GetControlText('QBS_VALAFFR'+IntToStr(i)) = 'ETP' then
        begin
          SetControlVisible('QBS_VALAFFR'+IntToStr(i)+'A_',false);
          SetControlVisible('QBS_VALAFFR'+IntToStr(i)+'A',false);
          SetControlVisible('QBS_VALAFFR'+IntToStr(i)+'B',false);
        end;
        if GetControlText('QBS_VALAFFR'+IntToStr(i)) = 'PSA' then
        begin
          LibellesRubSalPaie('QBS_VALAFFR'+IntToStr(i)+'A');
          SetControlVisible('QBS_VALAFFR'+IntToStr(i)+'A_',false);
          SetControlVisible('QBS_VALAFFR'+IntToStr(i)+'A',true);
        end;
      end;
    end;
  end;
  TabValAffSelected := false;
end;

procedure TOM_QBPSESSIONBP.ValRub_OnChange(Sender: TObject);
var ValAff,ValAff2:string;
begin
  ValAff := TControl( Sender ).name;
  ValAff2 := LeftStr(ValAff,length(ValAff)-2);
  if not TabValAffSelected then
  begin
    if GetControlText(ValAff) <> LastValue then
    begin
      if GetControlText(ValAff) = 'AAA' then
      begin
        SetControlText(ValAff2+'A','AAA');
        {$IFDEF EAGLCLIENT}
        THMultivalComboBox(GetControl(ValAff)).Aucun := true;
        THMultivalComboBox(GetControl(ValAff)).SelText := 'AAA';
        {$ELSE}
        THDBMultivalComboBox(GetControl(ValAff)).Aucun := true;
        THDBMultivalComboBox(GetControl(ValAff)).SelText := 'AAA';
        {$ENDIF}
        SetControlText(ValAff2+'B',TraduireMemoire('<<Tous>>'));
        SetControlProperty(ValAff2+'B','DATATYPE','PGREMUNERATION');
        SetControlVisible(ValAff2+'B',true);
      end;
      if GetControlText(ValAff) = 'COT' then
      begin
        SetControlText(ValAff2+'A','COT');
        SetControlText(ValAff2+'B',TraduireMemoire('<<Tous>>'));
        SetControlProperty(ValAff2+'B','DATATYPE','PGCOTIS');
        SetControlVisible(ValAff2+'B',true);
      end;
      if GetControlText(ValAff) = 'BAS' then
      begin
        SetControlText(ValAff2+'A','BAS');
        SetControlText(ValAff2+'B',TraduireMemoire('<<Tous>>'));
        SetControlProperty(ValAff2+'B','DATATYPE','PGBASECOTISATION');
        SetControlVisible(ValAff2+'B',true);
      end
    end;
  end;
end;

procedure TOM_QBPSESSIONBP.ValRub_OnEnter(Sender: TObject);
var ValAff:string;
begin
  ValAff := TControl( Sender ).name;
  LastValue := GetControlText(ValAff)
end;

procedure TOM_QBPSESSIONBP.TraitementRadioGroup;
var RadioGroup : THRadioGroup;
begin
  RadioGroup := THRadioGroup(GetControl('radiogrp'));
  if RadioGroup.itemindex = 0 then
  begin
    if ContextBP = 3 then
    begin
      SetControlVisible('btninit', FALSE);
      SetControlEnabled('btninit',FALSE);
      SetControlVisible('btninitg', TRUE);
      SetControlEnabled('btnloi', FALSE);
    end
    else
    begin
      SetControlVisible('btninit', TRUE);
      SetControlEnabled('btninit', TRUE);
      SetControlVisible('btninitg', FALSE);
      SetControlEnabled('btnloi', FALSE);
    end;
  end;
  if RadioGroup.itemindex = 1 then
  begin
    SetControlVisible('btninit', FALSE);
    SetControlVisible('btninitg', TRUE);
    { EVI / Empêche acces au menu loi si la session n'est pas initialisée }
    if THCheckBox(GetControl('QBS_SESSIONINIT')).Checked = TRUE then SetControlEnabled('btnloi', TRUE)
    else SetControlEnabled('btnloi', FALSE);
  end;
  //FQ Mode 15089
  SetControlEnabled('btnmenu',TRUE);
end;

{$IFDEF PAIEGRH}
procedure TOM_QBPSESSIONBP.LibellesLibrePaie(Champ:string;NbChamp:integer);
var Combo,i:integer;
begin
  for Combo := 1 to NbChamp do
  begin
    SetControlProperty(Champ+IntToStr(Combo),'DATATYPE','');
    for i := 1 to VH_Paie.PgNbreStatOrg do
    begin
      case i of
      1 : begin
            THValComboBox(GetControl(Champ+IntToStr(Combo))).Items.Add(VH_Paie.PgLibelleOrgStat1);
            THValComboBox(GetControl(Champ+IntToStr(Combo))).Values.Add('002');
          end;
      2 : begin
            THValComboBox(GetControl(Champ+IntToStr(Combo))).Items.Add(VH_Paie.PgLibelleOrgStat2);
            THValComboBox(GetControl(Champ+IntToStr(Combo))).Values.Add('003');
          end;
      3 : begin
            THValComboBox(GetControl(Champ+IntToStr(Combo))).Items.Add(VH_Paie.PgLibelleOrgStat3);
            THValComboBox(GetControl(Champ+IntToStr(Combo))).Values.Add('004');
          end;
      4 : begin
            THValComboBox(GetControl(Champ+IntToStr(Combo))).Items.Add(VH_Paie.PgLibelleOrgStat4);
            THValComboBox(GetControl(Champ+IntToStr(Combo))).Values.Add('005');
          end;
      end;
    end;

    for i := 1 to VH_Paie.PgNbCombo do
    begin
      case i of
      1 : begin
            THValComboBox(GetControl(Champ+IntToStr(Combo))).Items.Add(VH_Paie.PgLibCombo1);
            THValComboBox(GetControl(Champ+IntToStr(Combo))).Values.Add('006');
          end;
      2 : begin
            THValComboBox(GetControl(Champ+IntToStr(Combo))).Items.Add(VH_Paie.PgLibCombo2);
            THValComboBox(GetControl(Champ+IntToStr(Combo))).Values.Add('007');
          end;
      3 : begin
            THValComboBox(GetControl(Champ+IntToStr(Combo))).Items.Add(VH_Paie.PgLibCombo3);
            THValComboBox(GetControl(Champ+IntToStr(Combo))).Values.Add('008');
          end;
      4 : begin
            THValComboBox(GetControl(Champ+IntToStr(Combo))).Items.Add(VH_Paie.PgLibCombo4);
            THValComboBox(GetControl(Champ+IntToStr(Combo))).Values.Add('009');
          end;
      end;
    end;

    for i := 1 to VH_Paie.PGNbCoche do
    begin
      case i of
      1 : begin
            THValComboBox(GetControl(Champ+IntToStr(Combo))).Items.Add(VH_Paie.PgLibCoche1);
            THValComboBox(GetControl(Champ+IntToStr(Combo))).Values.Add('012');
          end;
      2 : begin
            THValComboBox(GetControl(Champ+IntToStr(Combo))).Items.Add(VH_Paie.PgLibCoche2);
            THValComboBox(GetControl(Champ+IntToStr(Combo))).Values.Add('013');
          end;
      3 : begin
            THValComboBox(GetControl(Champ+IntToStr(Combo))).Items.Add(VH_Paie.PgLibCoche3);
            THValComboBox(GetControl(Champ+IntToStr(Combo))).Values.Add('014');
          end;
      4 : begin
            THValComboBox(GetControl(Champ+IntToStr(Combo))).Items.Add(VH_Paie.PgLibCoche4);
            THValComboBox(GetControl(Champ+IntToStr(Combo))).Values.Add('015');
          end;
      end;
    end;

    if VH_Paie.PGLibCodeStat<>'' then
    begin
      THValComboBox(GetControl(Champ+IntToStr(Combo))).Items.Add(VH_Paie.PGLibCodeStat);
      THValComboBox(GetControl(Champ+IntToStr(Combo))).Values.Add('010');
    end;

  end;
end;

procedure TOM_QBPSESSIONBP.LibellesRubSalPaie(Champ:string);
var i:integer;
begin
  if Pos(Champ,ChampValAffOK)>0 then exit;
  for i := 1 to VH_Paie.PgNbSalLib do
  begin
    case i of
    1 : begin
          THMultivalComboBox(GetControl(Champ)).Items.Add(VH_Paie.PgSalLib1);
          THMultivalComboBox(GetControl(Champ)).Values.Add('SM1');
        end;
    2 : begin
          THMultivalComboBox(GetControl(Champ)).Items.Add(VH_Paie.PgSalLib2);
          THMultivalComboBox(GetControl(Champ)).Values.Add('SM2');
        end;
    3 : begin
          THMultivalComboBox(GetControl(Champ)).Items.Add(VH_Paie.PgSalLib3);
          THMultivalComboBox(GetControl(Champ)).Values.Add('SM3');
        end;
    4 : begin
          THMultivalComboBox(GetControl(Champ)).Items.Add(VH_Paie.PgSalLib4);
          THMultivalComboBox(GetControl(Champ)).Values.Add('SM4');
        end;
    5 : begin
          THMultivalComboBox(GetControl(Champ)).Items.Add(VH_Paie.PgSalLib5);
          THMultivalComboBox(GetControl(Champ)).Values.Add('SM5');
        end;
    end;
  end;
  ChampValAffOK:=ChampValAffOK+';'+Champ;
end;


procedure TOM_QBPSESSIONBP.MultiSocS_OnClick(Sender: TObject);
begin
  TFFiche(Ecran).Bouge(nbPost) ;
  AglLanceFiche('Q','QUFVBPMULTISOC','','','SESSION='+GetControlText('QBS_CODESESSION'));
end;



{$ENDIF PAIEGRH}



Initialization
  registerclasses ( [ TOM_QBPSESSIONBP ] ) ;
end.

