{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 05/12/2003
Modifié le ... :   /  /
Description .. : Edition scoring formation
Mots clefs ... : EManager,FORMATION
*****************************************************************
PT1 28/09/2007  V_7  FL  Adaptation cursus + accès assistant
PT2 05/10/2007  V_7  FL  Sélection millesimes en cours uniquement, tri décroissant des millesimes à l'impression 
                         et correction de bugs mémoire
PT3 18/10/2007  V_7  FL  Ajout de l'état de l'inscription et de la notion de DIF 
}

unit UtofPGEMEditionsFor;

interface
uses StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls,
  {$IFDEF EAGLCLIENT}
  eQRS1, UtilEAgl,
  {$ELSE}
  dbTables, QRS1, EdtREtat,
  {$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOF, ParamDat,  ParamSoc, HQry, UTob, HTB97, ed_tools,PGOutils2,entpaie,
  HStatus,P5DEF,PGOutilsFormation,LookUp;



type
  TOF_PGEMEDITREA = class(TOF)
  private
    TobEtat : Tob;
    Arg,TypeUtilisat : String;
    procedure EditColl;
    //procedure ExitEdit(Sender : TObject);
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad;override;
    procedure OnClose;override;
    procedure SalarieElipsisClick(Sender: TObject);
    procedure RespElipsisClick(Sender: TObject);
    procedure AccesResp(Sender : Tobject);
  end;

type
  TOF_PGEMEDITPREV = class(TOF)
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad;override;
    procedure OnClose;override;
    private
    TobEtat : Tob;
    Arg,TypeUtilisat : String;
    //procedure ExitEdit(Sender : TObject);
    procedure EditColl;
    procedure SalarieElipsisClick(Sender: TObject);
    procedure RespElipsisClick(Sender: TObject);
    procedure AccesResp(Sender : Tobject);
  end;

type
  TOF_PGEMEDITBILAN = class(TOF)
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad;override;
    procedure OnClose;override;
    private
    TobEtat : Tob;
    TypeUtilisat : String;
    procedure SalarieElipsisClick(Sender: TObject);
    procedure RespElipsisClick(Sender: TObject);
    procedure AccesResp(Sender : Tobject);
  end;




implementation
procedure TOF_PGEMEDITREA.RespElipsisClick(Sender: TObject);
var StFrom, StWhere: string;
begin
  StFrom := 'INTERIMAIRES';
  //PT1- Début
  //StWhere := StWhere + ' AND (PSI_INTERIMAIRE="'+V_PGI.UserSalarie+'"'+  
   If TypeUtilisat = 'R' then
    StWhere := ' AND (PSI_INTERIMAIRE="'+V_PGI.UserSalarie+'" '+
    'OR PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES where PGS_CODESERVICE IN '+
    '(SELECT PSO_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_SERVICESUP AND '+
    'PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))'
   Else
    StWhere := ' AND (PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PGS_CODESERVICE IN '+
    '(SELECT PSO_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_SERVICESUP AND '+
    'PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")))';
  //PT1 - Fin
  LookupList(THEdit(Sender), 'Liste des responsables', StFrom, 'PSI_INTERIMAIRE', 'PSI_LIBELLE,PSI_PRENOM', StWhere, 'PSI_INTERIMAIRE', TRUE, -1);
end;

procedure TOF_PGEMEDITREA.AccesResp(Sender : Tobject);
begin
     SetControlEnabled('PSE_RESPONSFOR',GetCheckBoxState('MULTINIVEAU')=CbChecked);
     SetControlEnabled('TPSE_RESPONSFOR',GetCheckBoxState('MULTINIVEAU')=CbChecked);
     SetControlEnabled('CSAUTRESP',GetCheckBoxState('MULTINIVEAU')=CbChecked);
     if GetCheckBoxState('MULTINIVEAU')<>CbChecked then
     begin
          SetControlChecked('CSAUTRESP',False);
          SetControlText('PSE_RESPONSFOR','');
     end;
end;

procedure TOF_PGEMEDITREA.SalarieElipsisClick(Sender: TObject);
var
  StFrom, StWhere: string;
begin
  StWhere := '(PSA_DATESORTIE<="' + UsDateTime(IDate1900) + '" OR PSA_DATESORTIE IS NULL OR PSA_DATESORTIE>="' + UsDateTime(V_PGI.DateEntree) + '")';
  StFrom := 'SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE';
  //PT1 - Début
{  if (GetCheckBoxState('MULTINIVEAU')<>CbChecked) then StWhere := StWhere + ' AND (PSE_RESPONSFOR="'+V_PGI.UserSalarie+'")'
  else StWhere := StWhere + ' AND (PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"'+
  ' OR PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES where PGS_CODESERVICE IN '+
  '(SELECT PSO_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_SERVICESUP AND'+
  ' PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))';}
  StWHere := StWHere + ' AND ' + AdaptByRespEmanager (TypeUtilisat,'PSE',V_PGI.UserSalarie,(GetCheckBoxState('MULTINIVEAU')=CbChecked));
  //PT1 - Fin
  LookupList(THEdit(Sender), 'Liste des salariés', StFrom, 'PSA_SALARIE', 'PSA_LIBELLE,PSA_PRENOM', StWhere, 'PSA_SALARIE', TRUE, -1);
end;

procedure TOF_PGEMEDITREA.OnLoad;
begin
     inherited;
   //PT1 - Début
{   If GetCheckBoxState('MULTINIVEAU') = CbChecked then
   begin
        SetControlText('XX_WHERE','(PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"'+
        ' OR PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES where PGS_CODESERVICE IN '+
        '(SELECT PSO_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_SERVICESUP AND'+
        ' PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))');
   end
   else
   begin
          SetControlText('XX_WHERE','PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"');
   end; }
   SetControlText('XX_WHERE',AdaptByRespEmanager (TypeUtilisat,'PSE',V_PGI.UserSalarie,(GetCheckBoxState('MULTINIVEAU')=CbChecked)));
   //PT1 - Fin
   If arg = 'COLLECTIF' then
   begin
        If GetCheckBoxState('CSAUTRESP') = CbChecked then
        begin
             TFQRs1(Ecran).CodeEtat := 'PRR';
             TFQRS1(Ecran).FCodeEtat := 'PRR';
        end
        else
        begin
             TFQRs1(Ecran).CodeEtat := 'PCR';
             TFQRS1(Ecran).FCodeEtat := 'PCR';
        end;
   end
   else
   begin
        If GetCheckBoxState('CSAUTRESP') = CbChecked then
        begin
             TFQRs1(Ecran).CodeEtat := 'PRA';
             TFQRS1(Ecran).FCodeEtat := 'PRA';
        end
        else
        begin
             TFQRs1(Ecran).CodeEtat := 'PIR';
             TFQRS1(Ecran).FCodeEtat := 'PIR';
        end
   end;
     If Arg = 'COLLECTIF' then
     begin
          If TobEtat<> Nil then FreeAndNil(TobEtat);
          TobEtat := Tob.Create('Edition',Nil,-1);
          EditColl;
          If GetCheckBoxState('CSAUTRESP') = CbChecked then TobEtat.Detail.Sort('PSE_RESPONSFOR');
          TFQRS1(Ecran).LaTob:= TobEtat;
     end;
end;

procedure TOF_PGEMEDITREA.OnClose;
begin
     inherited;
     If TobEtat<> Nil then FreeAndNil(TobEtat);
end;


procedure TOF_PGEMEDITREA.OnArgument(Arguments: string);
var
  Edit : THEdit;
  Q : TQuery;
  DD,DF : TDateTime;
  Check : TcheckBox;
begin
  inherited;
  //PT1 - Début
  If ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"') then TypeUtilisat := 'R'
  else TypeUtilisat := 'S';
  //PT1 - Fin
  Edit := THEdit(GetControl('PFO_SALARIE'));
  if Edit <> nil then Edit.OnElipsisClick := SalarieElipsisClick;
  Edit := THEdit(GetControl('PSE_RESPONSFOR'));
  if Edit <> nil then Edit.OnElipsisClick := RespElipsisClick;
  DD := V_PGI.DateEntree;
  DF := V_PGI.DateEntree;
  Q := openSQL('SELECT PFE_DATEDEBUT,PFE_DATEFIN FROM EXERFORMATION WHERE PFE_ACTIF="X" AND PFE_CLOTURE="-"',True);
  If Not Q.eof then
  begin
       DD := Q.FindField('PFE_DATEDEBUT').AsDateTime;
       DF := Q.FindField('PFE_DATEFIN').AsDateTime;
  end;
  Ferme(Q);
  SetControlText('PFO_DATEDEBUT',DateToStr(DD));
  SetControlText('PFO_DATEFIN',DateToStr(DF));
  Arg := ReadTokenPipe(Arguments, ';');
  If arg = 'COLLECTIF' then
   begin
        TFQRs1(Ecran).CodeEtat := 'PCR';
        TFQRS1(Ecran).FCodeEtat := 'PCR';
        TFQRS1(Ecran).Caption := 'Edition collective du plan réalisé';
   end
   else
   begin
        TFQRs1(Ecran).CodeEtat := 'PIR';
        TFQRS1(Ecran).FCodeEtat := 'PIR';
        TFQRS1(Ecran).Caption := 'Edition individuelle du plan réalisé';
   end;
   UpdateCaption(Ecran);
   Check := TcheckBox(getControl('MULTINIVEAU'));
   If Check <> Nil then Check.OnClick := AccesResp;
end;


procedure TOF_PGEMEDITREA.EditColl;
var
  Q: TQuery;
  TobRea,T : Tob;
  i : Integer;
  Where,LibEmploi,SQL : String;
  Pages : TPageControl;
  SQL1,SQL2,SQL3 : String;
  WhereResponsable : String;
begin
  Pages := TPageControl(GetControl('Pages'));
  Where := RecupWhereCritere(Pages);
  //PT1 - Début
  {If GetCheckBoxState('MULTINIVEAU') = CbChecked then
  begin
        WhereResponsable :=' AND (PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"'+
        ' OR PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES where PGS_CODESERVICE IN '+
        '(SELECT PSO_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_SERVICESUP AND'+
        ' PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))';
  end
  else WhereResponsable := ' AND PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"'; }
  WhereResponsable := ' AND '+AdaptByRespEmanager (TypeUtilisat,'PSE',V_PGI.UserSalarie,(GetCheckBoxState('MULTINIVEAU')=CbChecked));
  //PT1 - Fin
  If getCheckBoxState('CSAUTRESP') = CbChecked then SQL := 'SELECT PSE_RESPONSFOR,PFO_LIBEMPLOIFOR,COUNT (DISTINCT PFO_SALARIE) NBSAL,COUNT(PFO_CODESTAGE) NBSTAGE,SUM(PFO_NBREHEURE) NBH'+
  ',SUM(PFO_FRAISREEL) CTFRAIS,SUM (PFO_AUTRECOUT) CTPEDA,SUM(PFO_COUTREELSAL) CTSALAIRE FROM FORMATIONS LEFT JOIN DEPORTSAL ON PSE_SALARIE=PFO_SALARIE '+Where+
  ' GROUP BY PSE_RESPONSFOR,PFO_LIBEMPLOIFOR'
  else SQL := 'SELECT PFO_LIBEMPLOIFOR,COUNT (DISTINCT PFO_SALARIE) NBSAL,COUNT(PFO_CODESTAGE) NBSTAGE,SUM(PFO_NBREHEURE) NBH'+
  ',SUM(PFO_FRAISREEL) CTFRAIS,SUM (PFO_AUTRECOUT) CTPEDA,SUM(PFO_COUTREELSAL) CTSALAIRE FROM FORMATIONS LEFT JOIN DEPORTSAL ON PSE_SALARIE=PFO_SALARIE '+Where+
  ' GROUP BY PFO_LIBEMPLOIFOR';
  Q := OpenSQL(SQL,True);
  TobRea := Tob.Create('Lerealise', nil, -1);
  TobRea.LoadDetailDB('Lerealise', '', '', Q, False);
  Ferme(Q);
  for i := 0 to TobRea.detail.Count - 1 do
  begin
       T := Tob.Create('filleEdition',TobEtat,-1);
       LibEmploi := TobRea.Detail[i].GetValue('PFO_LIBEMPLOIFOR');
       T.AddChampSupValeur('METIER',RechDom('PGLIBEMPLOI',LibEmploi,False));
       T.AddChampSupValeur('NBSAL',TobRea.Detail[i].GetValue('NBSAL'));
       If getCheckBoxState('CSAUTRESP') = CbChecked then
       begin
            SQL1 := 'SELECT COUNT(PSA_SALARIE) EFFECTIF FROM SALARIES LEFT JOIN DEPORTSAL ON PSE_SALARIE=PSA_SALARIE WHERE PSE_RESPONSFOR="'+TobRea.Detail[i].GetValue('PSE_RESPONSFOR')+'"'+
       ' AND PSA_LIBELLEEMPLOI="'+TobRea.Detail[i].GetValue('PFO_LIBEMPLOIFOR')+'" AND (PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'" OR  PSA_DATESORTIE IS NULL OR PSA_DATESORTIE>="'+UsDateTime(V_PGI.DateEntree)+'")';
            SQL2 := 'SELECT COUNT(PFO_CODESTAGE) NBDIF FROM FORMATIONS LEFT JOIN DEPORTSAL ON PSE_SALARIE=PFO_SALARIE '+Where+' AND '+
            'PSE_RESPONSFOR="'+TobRea.Detail[i].GetValue('PSE_RESPONSFOR')+'" AND PFO_LIBEMPLOIFOR="'+LibEmploi+'" AND PFO_TYPEPLANPREV="DIF"';
            SQL3 := 'SELECT COUNT(PFI_CODESTAGE) NBDIFACCOR FROM INSCFORMATION LEFT JOIN DEPORTSAL ON PSE_SALARIE=PFI_SALARIE WHERE '+
            'PSE_RESPONSFOR="'+TobRea.Detail[i].GetValue('PSE_RESPONSFOR')+'" AND PFI_LIBEMPLOIFOR="'+LibEmploi+'" AND PFI_TYPEPLANPREV="DIF"';
            T.AddChampSupValeur('PSE_RESPONSFOR',TobRea.Detail[i].GetValue('PSE_RESPONSFOR'));
       end
       else
       begin
              SQL1 := 'SELECT COUNT(PSA_SALARIE) EFFECTIF FROM SALARIES LEFT JOIN DEPORTSAL ON PSE_SALARIE=PSA_SALARIE WHERE '+
       ' AND PSA_LIBELLEEMPLOI="'+TobRea.Detail[i].GetValue('PFO_LIBEMPLOIFOR')+'" AND (PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'" OR  PSA_DATESORTIE IS NULL OR PSA_DATESORTIE>="'+UsDateTime(V_PGI.DateEntree)+'")'+WhereResponsable;
            SQL2 := 'SELECT COUNT(PFO_CODESTAGE) NBDIF FROM FORMATIONS LEFT JOIN DEPORTSAL ON PSE_SALARIE=PFO_SALARIE '+Where+' AND PFO_LIBEMPLOIFOR="'+LibEmploi+'" AND PFO_TYPEPLANPREV="DIF"';
            SQL3 := 'SELECT COUNT(PFI_CODESTAGE) NBDIFACCOR FROM INSCFORMATION LEFT JOIN DEPORTSAL ON PSE_SALARIE=PFI_SALARIE WHERE PFI_LIBEMPLOIFOR="'+LibEmploi+'" AND PFI_TYPEPLANPREV="DIF"'+WhereResponsable ;
       end;
       Q := OpenSQL(SQL1,True);
       if Not Q.eof then T.AddChampSupValeur('EFFECTIF',Q.FindField('EFFECTIF').Asinteger)
       else T.AddChampSupValeur('EFFECTIF',0);
       ferme(Q);
       Q := OpenSQL(SQL2,True);
       if Not Q.eof then T.AddChampSupValeur('DIF',Q.FindField('NBDIF').Asinteger)
       else T.AddChampSupValeur('DIF',0);
       ferme(Q);
       Q := OpenSQL(SQL3,True);
       if Not Q.eof then T.AddChampSupValeur('DIFACCOR',Q.FindField('NBDIFACCOR').Asinteger)
       else T.AddChampSupValeur('DIFACCOR',0);
       ferme(Q);
       T.AddChampSupValeur('NBSTAGES',TobRea.Detail[i].GetValue('NBSTAGE'));
       T.AddChampSupValeur('NBHEURES',TobRea.Detail[i].GetValue('NBH'));
       T.AddChampSupValeur('COUT',TobRea.Detail[i].GetValue('CTFRAIS')+TobRea.Detail[i].GetValue('CTPEDA')+TobRea.Detail[i].GetValue('CTSALAIRE'));
       If GetCheckBoxState('CSAUTRESP') = CbChecked then T.AddChampSupValeur('PSE_RESPONSFOR',TobRea.Detail[i].GetValue('PSE_RESPONSFOR'));
  end;
  TobRea.Free;
end;



{TOF_PGEMEDITPREV}

procedure TOF_PGEMEDITPREV.RespElipsisClick(Sender: TObject);
var StFrom, StWhere: string;
begin
  StFrom := 'INTERIMAIRES';
  //PT1 - Début
{  StWhere := StWhere + ' AND (PSI_INTERIMAIRE="'+V_PGI.UserSalarie+'"'+
  ' OR PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES where PGS_CODESERVICE IN '+
  '(SELECT PSO_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_SERVICESUP AND'+
  ' PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))';}
   If TypeUtilisat = 'R' then
    StWhere := ' AND (PSI_INTERIMAIRE="'+V_PGI.UserSalarie+'" '+
    'OR PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES where PGS_CODESERVICE IN '+
    '(SELECT PSO_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_SERVICESUP AND '+
    'PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))'
   Else
    StWhere := ' AND (PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PGS_CODESERVICE IN '+
    '(SELECT PSO_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_SERVICESUP AND '+
    'PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")))';
  //PT1 -Fin
  LookupList(THEdit(Sender), 'Liste des responsables', StFrom, 'PSI_INTERIMAIRE', 'PSI_LIBELLE,PSI_PRENOM', StWhere, 'PSI_INTERIMAIRE', TRUE, -1);
end;


procedure TOF_PGEMEDITPREV.AccesResp(Sender : Tobject);
begin
     SetControlEnabled('PFI_RESPONSFOR',GetCheckBoxState('MULTINIVEAU')=CbChecked);
     SetControlEnabled('TPFI_RESPONSFOR',GetCheckBoxState('MULTINIVEAU')=CbChecked);
     SetControlEnabled('CSAUTRESP',GetCheckBoxState('MULTINIVEAU')=CbChecked);
     if GetCheckBoxState('MULTINIVEAU')<>CbChecked then
     begin
          SetControlChecked('CSAUTRESP',False);
          SetControlText('PFI_RESPONSFOR','');
     end;
end;


procedure TOF_PGEMEDITPREV.SalarieElipsisClick(Sender: TObject);
var
  StFrom, StWhere: string;
begin
  StWhere := '(PSA_DATESORTIE<="' + UsDateTime(IDate1900) + '" OR PSA_DATESORTIE IS NULL OR PSA_DATESORTIE>="' + UsDateTime(V_PGI.DateEntree) + '")';
  StFrom := 'SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE';
  //PT1 - Début
{  if GetCheckBoxState('MULTINIVEAU')<>CbChecked then StWhere := StWhere + ' AND (PSE_RESPONSFOR="'+V_PGI.UserSalarie+'")'
  else StWhere := StWhere + ' AND (PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"'+
  ' OR PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES where PGS_CODESERVICE IN '+
  '(SELECT PSO_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_SERVICESUP AND'+
  ' PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))';}
  StWhere := StWhere + ' AND ' + AdaptByRespEmanager (TypeUtilisat,'PSE',V_PGI.UserSalarie,(GetCheckBoxState('MULTINIVEAU')=CbChecked));
  //PT1 - Fin
  LookupList(THEdit(Sender), 'Liste des salariés', StFrom, 'PSA_SALARIE', 'PSA_LIBELLE,PSA_PRENOM', StWhere, 'PSA_SALARIE', TRUE, -1);
end;

procedure TOF_PGEMEDITPREV.OnArgument(Arguments: string);
var
    Edit : THEdit;
    Check : TcheckBox;
begin
   inherited;
  //PT1 - Début
  If ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"') then TypeUtilisat := 'R'
  else TypeUtilisat := 'S';
  //PT1 - Fin
   Arg := Arguments;
   SetControlText('PFI_MILLESIME',RendMillesimeEManager);
   If arg = 'SUIVIVALID' then
   begin
        TFQRs1(Ecran).CodeEtat := 'PSP';
        TFQRS1(Ecran).FCodeEtat := 'PSP';
        TFQRS1(Ecran).Caption := 'Suivi validation du prévisionnel';
        //PT2 - Début
        // Valeurs par défaut
        SetControlText('PFI_ETATINSCFOR', 'VAL');
        SetControlText('PFI_NATUREFORM', '002');
        //PT2 - Fin
   end
   else
   If arg = 'COLLECTIF' then
   begin
        TFQRs1(Ecran).CodeEtat := 'PCP';
        TFQRS1(Ecran).FCodeEtat := 'PCP';
        TFQRS1(Ecran).Caption := 'Edition collective du plan prévisionnel';
   end
   else
   begin
        TFQRs1(Ecran).CodeEtat := 'PIP';
        TFQRS1(Ecran).FCodeEtat := 'PIP';
        TFQRS1(Ecran).Caption := 'Edition individuelle du plan prévisionnel';
   end;
   UpdateCaption(Ecran);
   Edit := THEdit(GetControl('PFI_SALARIE'));
   if Edit <> nil then Edit.OnElipsisClick := SalarieElipsisClick;
   Edit := THEdit(GetControl('PFI_RESPONSFOR'));
  if Edit <> nil then Edit.OnElipsisClick := RespElipsisClick;
   Check := TcheckBox(getControl('MULTINIVEAU'));
   If Check <> Nil then Check.OnClick := AccesResp;
   SetControlProperty('PFI_MILLESIME','Plus','PFE_OUVREPREV="X"'); //PT2
end;


procedure TOF_PGEMEDITPREV.EditColl;
var
  Q: TQuery;
  TobRea,T : Tob;
  i : Integer;
  Pages : TPageControl;
  WhereResponsable,Where,LibEmploi,SQL : String;
begin
  Pages := TPageControl(GetControl('Pages'));
  Where := RecupWhereCritere(Pages);
  Where := Where + ' AND PFI_CODESTAGE<>"--CURSUS--" '; //PT1
  //PT1 - Début
{  If GetCheckBoxState('MULTINIVEAU') = CbChecked then
  begin
        WhereResponsable :=' AND (PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"'+
        ' OR PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES where PGS_CODESERVICE IN '+
        '(SELECT PSO_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_SERVICESUP AND'+
        ' PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))';
  end
  else WhereResponsable := ' AND PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"';}
  WhereResponsable := ' AND ' + AdaptByRespEmanager(TypeUtilisat,'PSE',V_PGI.UserSalarie,(GetCheckBoxState('MULTINIVEAU') = CbChecked));
  //PT1 - Fin
  If GetCheckBoxState('CSAUTRESP') = CbChecked then SQL := 'SELECT PFI_RESPONSFOR,PFI_LIBEMPLOIFOR,COUNT (DISTINCT PFI_SALARIE) NBSAL,SUM(PFI_NBINSC) NBSTAGE,SUM(PFI_DUREESTAGE) NBH'+
       ',SUM(PFI_FRAISFORFAIT) CTFRAIS,SUM (PFI_AUTRECOUT) CTPEDA,SUM(PFI_COUTREELSAL) CTSALAIRE FROM INSCFORMATION '+Where +
       ' GROUP BY PFI_RESPONSFOR,PFI_LIBEMPLOIFOR'
  else SQL := 'SELECT PFI_LIBEMPLOIFOR,COUNT (DISTINCT PFI_SALARIE) NBSAL,SUM(PFI_NBINSC) NBSTAGE,SUM(PFI_DUREESTAGE) NBH'+
       ',SUM(PFI_FRAISFORFAIT) CTFRAIS,SUM (PFI_AUTRECOUT) CTPEDA,SUM(PFI_COUTREELSAL) CTSALAIRE FROM INSCFORMATION '+Where +
       ' GROUP BY PFI_LIBEMPLOIFOR';
  Q := OpenSQL(SQL,True);
  TobRea := Tob.Create('Lerealise', nil, -1);
  TobRea.LoadDetailDB('Lerealise', '', '', Q, False);
  Ferme(Q);
  TobEtat := Tob.Create('Edition', nil, -1);
  for i := 0 to TobRea.detail.Count - 1 do
  begin
       T := Tob.Create('filleEdition',TobEtat,-1);
       LibEmploi := TobRea.Detail[i].GetValue('PFI_LIBEMPLOIFOR');
       T.AddChampSupValeur('METIER',RechDom('PGLIBEMPLOI',LibEmploi,False));
       T.AddChampSupValeur('NBSAL',TobRea.Detail[i].GetValue('NBSAL'));
       If GetCheckBoxState('CSAUTRESP') = CbChecked then SQL := 'SELECT COUNT(PSA_SALARIE) EFFECTIF FROM SALARIES LEFT JOIN DEPORTSAL ON PSE_SALARIE=PSA_SALARIE WHERE PSE_RESPONSFOR="'+TobRea.Detail[i].GetValue('PFI_RESPONSFOR')+'"'+
       ' AND PSA_LIBELLEEMPLOI="'+TobRea.Detail[i].GetValue('PFI_LIBEMPLOIFOR')+'" AND (PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'" OR  PSA_DATESORTIE IS NULL OR PSA_DATESORTIE>="'+UsDateTime(V_PGI.DateEntree)+'")'
       else SQL := 'SELECT COUNT(PSA_SALARIE) EFFECTIF FROM SALARIES LEFT JOIN DEPORTSAL ON PSE_SALARIE=PSA_SALARIE WHERE '+
       ' AND PSA_LIBELLEEMPLOI="'+TobRea.Detail[i].GetValue('PFI_LIBEMPLOIFOR')+'" AND (PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'" OR  PSA_DATESORTIE IS NULL OR PSA_DATESORTIE>="'+UsDateTime(V_PGI.DateEntree)+'")'+WhereResponsable;
       Q := OpenSQL(SQL,True);
       if Not Q.eof then T.AddChampSupValeur('EFFECTIF',Q.FindField('EFFECTIF').Asinteger)
       else T.AddChampSupValeur('EFFECTIF',0);
       ferme(Q);
       If GetCheckBoxState('CSAUTRESP') = CbChecked then SQL := 'SELECT COUNT(PFO_CODESTAGE) NBDIF FROM FORMATIONS LEFT JOIN DEPORTSAL ON PSE_SALARIE=PFO_SALARIE WHERE PSE_RESPONSFOR="'+TobRea.Detail[i].GetValue('PFI_RESPONSFOR')+'" AND PFO_LIBEMPLOIFOR="'+LibEmploi+'" '+
       'AND PFO_TYPEPLANPREV="DIF"'
       else SQL := 'SELECT COUNT(PFO_CODESTAGE) NBDIF FROM FORMATIONS LEFT JOIN DEPORTSAL ON PSE_SALARIE=PFO_SALARIE WHERE PFO_LIBEMPLOIFOR="'+LibEmploi+'" AND PFO_TYPEPLANPREV="DIF"'+WhereResponsable;
       Q := OpenSQL(SQL,True);
       if Not Q.eof then T.AddChampSupValeur('DIF',Q.FindField('NBDIF').Asinteger)
       else T.AddChampSupValeur('DIF',0);
       ferme(Q);
       If GetCheckBoxState('CSAUTRESP') = CbChecked then SQL := 'SELECT COUNT(PFI_CODESTAGE) NBDIFACCOR FROM INSCFORMATION LEFT JOIN DEPORTSAL ON PSE_SALARIE=PFI_SALARIE '+Where+' AND '+
       'PSE_RESPONSFOR="'+TobRea.Detail[i].GetValue('PFI_RESPONSFOR')+'" AND PFI_LIBEMPLOIFOR="'+LibEmploi+'" AND PFI_TYPEPLANPREV="DIF"'
       else SQL := 'SELECT COUNT(PFI_CODESTAGE) NBDIFACCOR FROM INSCFORMATION LEFT JOIN DEPORTSAL ON PSE_SALARIE=PFI_SALARIE '+Where+' AND PFI_LIBEMPLOIFOR="'+LibEmploi+'" AND PFI_TYPEPLANPREV="DIF"';
       Q := OpenSQL(SQL,True);
       if Not Q.eof then T.AddChampSupValeur('DIFACCOR',Q.FindField('NBDIFACCOR').Asinteger)
       else T.AddChampSupValeur('DIFACCOR',0);
       ferme(Q);
       T.AddChampSupValeur('NBSTAGES',TobRea.Detail[i].GetValue('NBSTAGE'));
       T.AddChampSupValeur('NBHEURES',TobRea.Detail[i].GetValue('NBH'));
       T.AddChampSupValeur('COUT',TobRea.Detail[i].GetValue('CTFRAIS')+TobRea.Detail[i].GetValue('CTPEDA')+TobRea.Detail[i].GetValue('CTSALAIRE'));
       T.AddChampSupValeur('HONORAIRE',TobRea.Detail[i].GetValue('CTPEDA'));
       If GetCheckBoxState('CSAUTRESP') = CbChecked then T.AddChampSupValeur('PFI_RESPONSFOR',TobRea.Detail[i].GetValue('PFI_RESPONSFOR'));
  end;
  TobRea.Free;
end;

procedure TOF_PGEMEDITPREV.OnLoad;
begin
     inherited;
   //PT1 - Début
{   If GetCheckBoxState('MULTINIVEAU') = CbChecked then
   begin
        SetControlText('XX_WHERE','(PFI_RESPONSFOR="'+V_PGI.UserSalarie+'"'+
        ' OR PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES where PGS_CODESERVICE IN '+
        '(SELECT PSO_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_SERVICESUP AND'+
        ' PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))');
   end
   else
   begin
          SetControlText('XX_WHERE','PFI_RESPONSFOR="'+V_PGI.UserSalarie+'"');
   end;}
   SetControlText('XX_WHERE',AdaptByRespEmanager (TypeUtilisat,'PFI',V_PGI.UserSalarie,(GetCheckBoxState('MULTINIVEAU')=CbChecked)));
   //PT1 - Fin
   If arg = 'SUIVIVALID' then
   begin
        If GetCheckBoxState('CSAUTRESP') = CbChecked then
        begin
             TFQRs1(Ecran).CodeEtat := 'PSR';
             TFQRS1(Ecran).FCodeEtat := 'PSR';
        end
        else
        begin
             TFQRs1(Ecran).CodeEtat := 'PSP';
             TFQRS1(Ecran).FCodeEtat := 'PSP';
        end;
   end
   else
   If arg = 'COLLECTIF' then
   begin
        If GetCheckBoxState('CSAUTRESP') = CbChecked then
        begin
             TFQRs1(Ecran).CodeEtat := 'PRC';
             TFQRS1(Ecran).FCodeEtat := 'PRC';
        end
        else
        begin
             TFQRs1(Ecran).CodeEtat := 'PCP';
             TFQRS1(Ecran).FCodeEtat := 'PCP';
        end;
   end
   else
   begin
        If GetCheckBoxState('CSAUTRESP') = CbChecked then
        begin
             TFQRs1(Ecran).CodeEtat := 'PRI';
             TFQRS1(Ecran).FCodeEtat := 'PRI';
        end
        else
        begin
             TFQRs1(Ecran).CodeEtat := 'PIP';
             TFQRS1(Ecran).FCodeEtat := 'PIP';
        end;
   end;
     If Arg = 'COLLECTIF' then
     begin
          If TobEtat<> Nil then FreeAndNil(TobEtat);
//          TobEtat := Tob.Create('Edition',Nil,-1); //PT2
          EditColl;
          If GetCheckBoxState('CSAUTRESP') = CbChecked then TobEtat.Detail.Sort('PFI_RESPONSFOR');
          TFQRS1(Ecran).LaTob:= TobEtat;
     end;
end;

procedure TOF_PGEMEDITPREV.OnClose;
begin
     inherited;
     If TobEtat<> Nil then FreeAndNil(TobEtat);
end;

{TOF_PGEMEDITBILAN}
procedure TOF_PGEMEDITBILAN.RespElipsisClick(Sender: TObject);
var StFrom, StWhere: string;
begin
  StFrom := 'INTERIMAIRES';
  //PT1 - Début
{  StWhere := StWhere + ' AND (PSI_INTERIMAIRE="'+V_PGI.UserSalarie+'"'+
  ' OR PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES where PGS_CODESERVICE IN '+
  '(SELECT PSO_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_SERVICESUP AND'+
  ' PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))';}
   If TypeUtilisat = 'R' then
    StWhere := ' AND (PSI_INTERIMAIRE="'+V_PGI.UserSalarie+'" '+
    'OR PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES where PGS_CODESERVICE IN '+
    '(SELECT PSO_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_SERVICESUP AND '+
    'PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))'
   Else
    StWhere := ' AND (PSI_INTERIMAIRE IN (SELECT PGS_RESPONSFOR FROM SERVICES WHERE PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'" OR PGS_CODESERVICE IN '+
    '(SELECT PSO_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_SERVICESUP AND '+
    'PGS_SECRETAIREFOR="'+V_PGI.UserSalarie+'")))';
  //PT1 - Fin
  LookupList(THEdit(Sender), 'Liste des responsables', StFrom, 'PSI_INTERIMAIRE', 'PSI_LIBELLE,PSI_PRENOM', StWhere, 'PSI_INTERIMAIRE', TRUE, -1);
end;


procedure TOF_PGEMEDITBILAN.AccesResp(Sender : Tobject);
begin
     SetControlEnabled('PFI_RESPONSFOR',GetCheckBoxState('MULTINIVEAU')=CbChecked);
     SetControlEnabled('TPFI_RESPONSFOR',GetCheckBoxState('MULTINIVEAU')=CbChecked);
end;


procedure TOF_PGEMEDITBILAN.SalarieElipsisClick(Sender: TObject);
var
  StFrom, StWhere: string;
begin
  StWhere := '(PSA_DATESORTIE<="' + UsDateTime(IDate1900) + '" OR PSA_DATESORTIE IS NULL OR PSA_DATESORTIE>="' + UsDateTime(V_PGI.DateEntree) + '")';
  StFrom := 'SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE';
  //PT1 - Début
  {if GetCheckBoxState('MULTINIVEAU')<>CbChecked then StWhere := StWhere + ' AND (PSE_RESPONSFOR="'+V_PGI.UserSalarie+'")'
  else StWhere := StWhere + ' AND (PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"'+
  ' OR PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES where PGS_CODESERVICE IN '+
  '(SELECT PSO_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_SERVICESUP AND'+
  ' PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))';}
  StWHere := StWHere + ' AND ' + AdaptByRespEmanager (TypeUtilisat,'PSE',V_PGI.UserSalarie,(GetCheckBoxState('MULTINIVEAU')=CbChecked));
  //PT1 - Fin
  LookupList(THEdit(Sender), 'Liste des salariés', StFrom, 'PSA_SALARIE', 'PSA_LIBELLE,PSA_PRENOM', StWhere, 'PSA_SALARIE', TRUE, -1);
end;

procedure TOF_PGEMEDITBILAN.OnArgument(Arguments: string);
var Edit : THEdit;
    Q : TQuery;
    Check : TcheckBox;
    Values : String; //PT2
begin
     inherited;
  //PT1 - Début
  If ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"') then TypeUtilisat := 'R'
  else TypeUtilisat := 'S';
  //PT1 - Fin
     Edit := THEdit(GetControl('SALARIE'));
     if Edit <> nil then Edit.OnElipsisClick := SalarieElipsisClick;
     Edit := THEdit(GetControl('PFI_RESPONSFOR'));
     if Edit <> nil then Edit.OnElipsisClick := RespElipsisClick;
     TFQRS1(Ecran).Caption := 'Edition du bilan individuel de formation';
     UpdateCaption(Ecran);
     //PT2 - Début
     //SetControlText('PFI_MILLESIME',RendMillesimeEManager);
     SetControlProperty('PFI_MILLESIME','Plus','PFE_OUVREPREV="X" OR (PFE_ENCOURS="X" AND PFE_OUVREPREV="-")');  // Uniquement les millesimes actifs
     // Préselection des deux derniers exercices par défaut
     Q := OpenSQL('SELECT PFE_MILLESIME FROM EXERFORMATION WHERE PFE_OUVREPREV="X" OR (PFE_ENCOURS="X" AND PFE_OUVREPREV="-") ORDER BY PFE_MILLESIME DESC', True);
     While Not Q.EOF Do
     Begin
          Values := Values + Q.FindField('PFE_MILLESIME').AsString + ';';
          If Length(Values) > 5 Then Break; // On limite à 2 exercices (4 car + ';')
          Q.Next;
     End;
     Ferme(Q);
     SetControlText('PFI_MILLESIME', Values);
     Q := OpenSQL('SELECT PFE_DATEDEBUT,PFE_DATEFIN FROM EXERFORMATION WHERE PFE_ACTIF="X" AND PFE_CLOTURE="-" ORDER BY PFE_DATEDEBUT DESC',True);
     //PT2 - Fin
     If Not Q.Eof then
     begin
          SetControlText('DATEDEB',DateToStr(Q.FindField('PFE_DATEDEBUT').AsdateTime));
          SetControlText('DATEFIN',DateToStr(Q.FindField('PFE_DATEFIN').AsdateTime));
     end;
     Ferme(Q);
     SetControlProperty('DATEFIN','DisplayFormat','');
     Check := TcheckBox(getControl('MULTINIVEAU'));
   If Check <> Nil then Check.OnClick := AccesResp;
   //PT1 - Début
   SetControlProperty('LIBEMPLOI','Plus', ' AND CC_CODE IN (SELECT PSA_LIBELLEEMPLOI FROM SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE '+
  'WHERE '+AdaptByRespEmanager (TypeUtilisat,'PSE',V_PGI.UserSalarie,True)+')');
  {'(PSE_RESPONSFOR="'+V_PGI.UserSalarie+'" OR PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
  ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="'+V_PGI.UserSalarie+'"))))');}
  //PT1 - Fin
end;

procedure TOF_PGEMEDITBILAN.OnLoad;
var Q : TQuery;
    TobRea,TobPrev,T : Tob;
    i : Integer;
    WhereRea,WherePrev,St : String; //PT1
    Pages : TPageControl;
    SDateD,SDateF : String;
    Millesime : String;
    Year, Month, Day, Year2, Month2, Day2 : Word; //PT1
    DD,DF : TDateTime; //PT1
begin
     inherited;
     Millesime := GetControlText('PFI_MILLESIME');
     SDateD := GetcontrolText('DATEDEB');
     SDateF := GetcontrolText('DATEFIN');
     Pages := TPageControl(GetControl('Pages'));
     WherePrev := RecupWhereCritere(Pages);
     WhereRea := 'WHERE PFO_EFFECTUE="X" AND PFO_DATEDEBUT>="'+UsdateTime(StrToDate(GetcontrolText('DATEDEB')))+'" AND PFO_DATEDEBUT<="'+UsdateTime(StrToDate(GetcontrolText('DATEFIN')))+'"';
     if GetControlText('SALARIE') <> '' then
     begin
          If WherePrev <> '' then WherePrev := WherePrev+' AND PFI_SALARIE="'+GetControltext('SALARIE')+'"'
          else WherePrev := 'WHERE PFI_SALARIE="'+GetControltext('SALARIE')+'"';
          Whererea := Whererea+' AND PFO_SALARIE="'+GetControltext('SALARIE')+'"'
     end;
     if GetControlText('LIBEMPLOI') <> '' then
     begin
          If WherePrev <> '' then WherePrev := WherePrev+' AND PSA_LIBELLEEMPLOI="'+GetControltext('LIBEMPLOI')+'"'
          else WherePrev := 'WHERE PSA_LIBELLEEMPLOI="'+GetControltext('LIBEMPLOI')+'"';
          Whererea := Whererea+' AND PSA_LIBELLEEMPLOI="'+GetControltext('LIBEMPLOI')+'"'
     end;
     If WherePrev <> '' then WherePrev := WherePrev+' AND PFI_SALARIE<>""'
     else WherePrev := 'WHERE PFI_SALARIE<>""';
     If GetCheckBoxState('MULTINIVEAU') = CbChecked then
     begin
          //PT1 - Début
          {WherePrev := WherePrev + ' AND (PFI_RESPONSFOR="'+V_PGI.UserSalarie+'"'+
          ' OR PFI_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES where PGS_CODESERVICE IN '+
          '(SELECT PSO_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_SERVICESUP AND'+
          ' PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))';
          WhereRea := WhereRea + ' AND (PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"'+
          ' OR PSE_RESPONSFOR IN (SELECT PGS_RESPONSFOR FROM SERVICES where PGS_CODESERVICE IN '+
          '(SELECT PSO_CODESERVICE FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_SERVICESUP AND'+
          ' PGS_RESPONSFOR="'+V_PGI.UserSalarie+'")))';}
          WherePrev := WherePrev + ' AND ' + AdaptByRespEmanager (TypeUtilisat,'PFI',V_PGI.UserSalarie,True);
          WhereRea  := WhereRea  + ' AND ' + AdaptByRespEmanager (TypeUtilisat,'PSE',V_PGI.UserSalarie,True);
          //PT1 - Fin
          If GetControlText('PFI_RESPONSFOR') <> '' then WhereRea := WhereRea + ' AND (PSE_RESPONSFOR="'+GetControlText('PFI_RESPONSFOR')+'"';
     end
     else
     begin
          //PT1 - Début
          {WherePrev := WherePrev + ' AND PFI_RESPONSFOR="'+V_PGI.UserSalarie+'"';
          WhereRea := WhereRea + ' AND PSE_RESPONSFOR="'+V_PGI.UserSalarie+'"';}
          WherePrev := WherePrev + ' AND ' + AdaptByRespEmanager (TypeUtilisat,'PFI',V_PGI.UserSalarie,False);
          WhereRea  := WhereRea  + ' AND ' + AdaptByRespEmanager (TypeUtilisat,'PSE',V_PGI.UserSalarie,False);
          //PT1 - Fin
     end;
     //PT1 - Début
     //Gestion de la date d'entrée
     If (GetControlText('DATEENTREE') <> '') And (GetControlText('DATEENTREE') <> '  /  /    ')  Then
     Begin
          WherePrev := WherePrev + ' AND PFI';
          WhereRea  := WhereRea  + ' AND PFO';

          St := '_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_DATEENTREE>="'+UsDateTime(StrToDate(GetControlText('DATEENTREE')))+'"';
          If (GetControlText('DATEENTREE_') <> '') And (GetControlText('DATEENTREE_') <> '  /  /    ') Then St := St + ' AND PSA_DATEENTREE<="'+UsDateTime(StrToDate(GetControlText('DATEENTREE_')))+'"';
          St := St + ')';

          WherePrev := WherePrev  + St;
          WhereRea := WhereRea  + St;
     End;
     //PT1 - Fin
     // Prévisionnel
     If TobEtat<> Nil then FreeAndNil(TobEtat);
     TobEtat := Tob.Create('Edition',Nil,-1);
     Q := OpenSQL('SELECT PSE_RESPONSFOR,PFI_SALARIE,PSA_LIBELLE,PSA_PRENOM,PFI_MILLESIME,PFI_CODESTAGE,PFI_DUREESTAGE,PFI_ETATINSCFOR,'+ //PT2 //PT3
      'PFI_NATUREFORM,PFI_TYPEPLANPREV,PSA_LIBELLEEMPLOI,PSA_DATEENTREE,PSA_DATENAISSANCE,PSA_ETABLISSEMENT,PFI_REALISE FROM INSCFORMATION '+
      'LEFT JOIN SALARIES ON PSA_SALARIE=PFI_SALARIE LEFT JOIN DEPORTSAL ON PSE_SALARIE=PFI_SALARIE '+WherePrev+' AND PFI_CODESTAGE<>"--CURSUS--" ORDER BY PFI_SALARIE ASC,PFI_MILLESIME DESC',True); //PT1 //PT2
      TobPrev := Tob.Create('leprevisionnel', nil, -1);
      TobPrev.LoadDetailDB('leprevisionnel', '', '', Q, False);
      Ferme(Q);
      for i := 0 to TobPrev.detail.Count - 1 do
      begin
       T := Tob.Create('filleEdition',TobEtat,-1);
       T.AddChampSupValeur('PFO_SALARIE',TobPrev.Detail[i].GetValue('PFI_SALARIE'));
       T.AddChampSupValeur('PFO_MILLESIME', TobPrev.Detail[i].GetValue('PFI_MILLESIME')); //PT1 //PT2
       T.AddChampSupValeur('PFO_TYPEPLANPREV','PREV');
       T.AddChampSupValeur('DATEFOR',IDate1900);
       T.AddChampSupValeur('LIBSALARIE',TobPrev.Detail[i].GetValue('PSA_LIBELLE') + ' '+ TobPrev.Detail[i].GetValue('PSA_PRENOM'));
       T.AddChampSupValeur('LIBSTAGE',RechDom('PGSTAGEFORMCOMPLET',TobPrev.Detail[i].GetValue('PFI_CODESTAGE'),False));
       T.AddChampSupValeur('LIBELLETYPE','Prévisionnel '+TobPrev.Detail[i].GetValue('PFI_MILLESIME')); //PT2
       T.AddChampSupValeur('NBHEURES',TobPrev.Detail[i].GetValue('PFI_DUREESTAGE'));
       T.AddChampSupValeur('NATURE',RechDom('PGNATUREFORM',TobPrev.Detail[i].GetValue('PFI_NATUREFORM'),False));
       T.AddChampSupValeur('LIBEMPLOI',RechDom('PGLIBEMPLOI',TobPrev.Detail[i].GetValue('PSA_LIBELLEEMPLOI'),False));
       T.AddChampSupValeur('DATEENTREE',TobPrev.Detail[i].GetValue('PSA_DATEENTREE'));
       T.AddChampSupValeur('DATENAISSANCE',TobPrev.Detail[i].GetValue('PSA_DATENAISSANCE'));
       T.AddChampSupValeur('ETABLISSEMENT',RechDom('TTETABLISSEMENT',TobPrev.Detail[i].GetValue('PSA_ETABLISSEMENT'),False));
       T.AddChampSupValeur('RESPONSABLE',RechDom('PGINTERIMAIRES',TobPrev.Detail[i].GetValue('PSE_RESPONSFOR'),False));
       T.AddChampSupValeur('TYPEPLANPREV',TobPrev.Detail[i].GetValue('PFI_TYPEPLANPREV')); //PT3
       T.AddChampSupValeur('ETATINSC',RechDom('PGETATVALIDATION',TobPrev.Detail[i].GetValue('PFI_ETATINSCFOR'),False)); //PT3
       If TobPrev.Detail[i].GetValue('PFI_REALISE') = 'X' then T.AddChampSupValeur('ETAT','Oui')
       else T.AddChampSupValeur('ETAT','Non');
      end;
      TobPrev.Free;
      // Réalisé
      Q := OpenSQL('SELECT PFO_DATEDEBUT,PSE_RESPONSFOR,PFO_SALARIE,PSA_LIBELLE,PSA_PRENOM,PFO_LIBEMPLOIFOR,PFO_CODESTAGE,PFO_NBREHEURE,PFO_TYPEPLANPREV,PFO_ETATINSCFOR,'+ //PT3
      'PFO_NATUREFORM,PFO_TYPEPLANPREV,PSA_LIBELLEEMPLOI,PSA_DATEENTREE,PSA_DATENAISSANCE,PSA_ETABLISSEMENT FROM FORMATIONS '+
      'LEFT JOIN SALARIES ON PSA_SALARIE=PFO_SALARIE LEFT JOIN DEPORTSAL ON PSE_SALARIE=PFO_SALARIE '+WhereRea+' ORDER BY PFO_DATEDEBUT DESC',True);  //PT2
      TobRea := Tob.Create('Lerealise', nil, -1);
      TobRea.LoadDetailDB('Lerealise', '', '', Q, False);
      Ferme(Q);
      for i := 0 to TobRea.detail.Count - 1 do
      begin
       T := Tob.Create('filleEdition',TobEtat,-1);
       T.AddChampSupValeur('PFO_SALARIE',TobRea.Detail[i].GetValue('PFO_SALARIE'));
       //PT1 - Début
       DecodeDate(TobRea.Detail[i].GetValue('PFO_DATEDEBUT'), Year, Month, Day);
       T.AddChampSupValeur('PFO_MILLESIME', Year);
       DecodeDate(StrToDate(sDateD), Year2, Month2, Day2);DD := EncodeDate(Year, Month2, Day2);
       DecodeDate(StrToDate(sDateF), Year2, Month2, Day2);DF := EncodeDate(Year, Month2, Day2);
       //PT1 - Fin
       T.AddChampSupValeur('PFO_TYPEPLANPREV','REA');
       T.AddChampSupValeur('DATEFOR',TobRea.Detail[i].GetValue('PFO_DATEDEBUT'));
       T.AddChampSupValeur('LIBSALARIE',TobRea.Detail[i].GetValue('PSA_LIBELLE') +' '+TobRea.Detail[i].GetValue('PSA_PRENOM'));
       T.AddChampSupValeur('LIBSTAGE',RechDom('PGSTAGEFORMCOMPLET',TobRea.Detail[i].GetValue('PFO_CODESTAGE'),False));
       T.AddChampSupValeur('LIBELLETYPE','Réalisé du '+DateToStr(DD)+ ' au '+DateToStr(DF)); //PT1
       T.AddChampSupValeur('NBHEURES',TobRea.Detail[i].GetValue('PFO_NBREHEURE'));
       T.AddChampSupValeur('NATURE',RechDom('PGNATUREFORM',TobRea.Detail[i].GetValue('PFO_NATUREFORM'),False));
       T.AddChampSupValeur('LIBEMPLOI',RechDom('PGLIBEMPLOI',TobRea.Detail[i].GetValue('PSA_LIBELLEEMPLOI'),False));
       T.AddChampSupValeur('DATEENTREE',TobRea.Detail[i].GetValue('PSA_DATEENTREE'));
       T.AddChampSupValeur('DATENAISSANCE',TobRea.Detail[i].GetValue('PSA_DATENAISSANCE'));
       T.AddChampSupValeur('ETABLISSEMENT',RechDom('TTETABLISSEMENT',TobRea.Detail[i].GetValue('PSA_ETABLISSEMENT'),False));
       T.AddChampSupValeur('RESPONSABLE',RechDom('PGINTERIMAIRES',TobRea.Detail[i].GetValue('PSE_RESPONSFOR'),False));
       T.AddChampSupValeur('TYPEPLANPREV',TobRea.Detail[i].GetValue('PFO_TYPEPLANPREV')); //PT3
       T.AddChampSupValeur('ETATINSC',TobRea.Detail[i].GetValue('PFO_ETATINSCFOR')); //PT3
       T.AddChampSupValeur('ETATINSC',RechDom('PGETATVALIDATION',TobRea.Detail[i].GetValue('PFO_ETATINSCFOR'),False)); //PT3
       If ExisteSQL('SELECT PSE_RESPONSFOR,PFI_SALARIE,PSA_LIBELLE,PSA_PRENOM,PFI_CODESTAGE,PFI_DUREESTAGE,'+
      'PFI_NATUREFORM,PFI_TYPEPLANPREV,PSA_LIBELLEEMPLOI,PSA_DATEENTREE,PSA_DATENAISSANCE,PSA_ETABLISSEMENT,PFI_REALISE FROM INSCFORMATION '+
      'LEFT JOIN SALARIES ON PSA_SALARIE=PFI_SALARIE LEFT JOIN DEPORTSAL ON PSE_SALARIE=PFI_SALARIE '+WherePrev+
      ' AND PFI_CODESTAGE="'+TobRea.Detail[i].GetValue('PFO_CODESTAGE')+'" AND PFI_SALARIE="'+TobRea.Detail[i].GetValue('PFO_SALARIE')+'"') then T.AddChampSupValeur('ETAT','Oui')
      else T.AddChampSupValeur('ETAT','Non');
      end;
      TobRea.Free;
      TobEtat.Detail.Sort('PFO_SALARIE;PFO_MILLESIME;PFO_TYPEPLANPREV;DATEFOR'); //PT1
      TFQRS1(Ecran).LaTob:= TobEtat;
end;

procedure TOF_PGEMEDITBILAN.OnClose;
begin
     inherited;
     If TobEtat<> Nil then FreeAndNil(TobEtat);
end;



initialization
  registerclasses([TOF_PGEMEDITREA,TOF_PGEMEDITPREV,TOF_PGEMEDITBILAN]);
end.

