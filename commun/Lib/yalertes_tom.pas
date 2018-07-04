{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 08/11/2004
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : YALERTES (YALERTES)
Mots clefs ... : TOM;YALERTES
*****************************************************************}
Unit YALERTes_tom ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db, DbCtrls,
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
     Fiche,
     FichList,
     FE_Main,
{$else}
     eFiche,
     eFichList,
     MaineAgl,UtileAGL,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,UTOM,
     UTob,extctrls,htb97,UtilAlertes,HRichOLE,HPanel,Hdb,buttons,
     ParamSoc,HSysMenu ;

Const
   stComboUtilisateur : String = 'TTUTILISATEUR';
   stComboGroupe: String = 'TTUSERGROUPE';
   stComboEvent: String = 'YEVENEMENTS';
   stNomChpEvent: String = 'YAL_EVENEMENT';
   stInitChp1: String = 'YAD_OPER';
   stInitChp2: String = 'YAD_VAL';
   Prefixe45 : String = 'GP'; { tables qui gère les evenements 4 et 5 }
   TesteMessage : array[1..7] of string  = (
        {1}         'La zone "Alerte sur" doit être renseignée',
        {2}         'Vous devez sélectionner un ou plusieurs utilisateur(s)',
        {3}         'Vous devez sélectionner un ou plusieurs groupe(s)',
        {4}         'Vous devez cocher un évènement',
        {5}         'L''alerte ouverture fiche ne peut pas être bloquante ou interrogative',
        {6}         'La saisie d''un code commencant par CEG est interdite',
        {7}         'L''alerte modification de champs ne peut pas être bloquante ou interrogative s''il y a des champs infos complémentaires'        
              );
Type
  TOM_YALERTES = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
  Private
    bDuplication,bMultiLine : boolean;
    
    stAction,oldCodeAlerte : string;
    NbEvent : Integer;
    TobDesti : Tob;
    CBCHAMP1 : THValcomboBox;
    CBCHAMP2 : THValcomboBox;
    CBCHAMP3 : THValcomboBox;
    CBCHAMP1_ : THValcomboBox;
    CBCHAMP2_ : THValcomboBox;
    CBCHAMP3_ : THValcomboBox;
    CBCHAMP1__ : THValcomboBox;
    CBCHAMP2__ : THValcomboBox;
    CBCHAMP3__ : THValcomboBox;

    ChampsDispo,ChampsCondi: HTStringList;
    LibellesDispo,LibellesCondi: HTStrings;
    ChpsDispo,ChpsCondi : THListBox;
    stArgument : boolean;
    InitChamp,InitValeur,InitLibelle : String;
    procedure ALCHAMP1Click (Sender: TObject);
    procedure ALCHAMP2Click (Sender: TObject);
    procedure ALCHAMP3Click (Sender: TObject);
    procedure ALCHAMP1_Click (Sender: TObject);
    procedure ALCHAMP2_Click (Sender: TObject);
    procedure ALCHAMP3_Click (Sender: TObject);
    procedure ALCHAMP1__Click (Sender: TObject);
    procedure ALCHAMP2__Click (Sender: TObject);
    procedure ALCHAMP3__Click (Sender: TObject);

    procedure ZoneClick (Sender: TObject);

    procedure AlimTobDestinataires;
    procedure ChgtDestinataires ;
    procedure ChargeListUser;
    procedure MajDestinataires(delete : boolean);
    procedure DestiClick(Sender: TObject);
    procedure TypeClick(Sender: TObject);
    procedure ChargeListeConditions(Prefixe,Suffixe : String);
    function  AffecteDatatype(CHAMPTABLE,CHAMPCONF: string):string;
    procedure chargeComboChamps( stPrefixe,ListeDhControle,TypeChamp : string;ListeChamp,ListeLibelle: HTStrings);
    procedure MajConditions(delete : boolean);
    procedure MajChamps (TypeChamp : String);
    procedure Evenement1Click(Sender: TObject);
    procedure Evenement2Click(Sender: TObject);
    procedure Evenement3Click(Sender: TObject);
    procedure Evenement4Click(Sender: TObject);
    procedure Evenement5Click(Sender: TObject);
    procedure Evenement6Click(Sender: TObject);
    procedure Evenement7Click(Sender: TObject);
    procedure Evenement8Click(Sender: TObject);

    procedure AlimListesChamps(tous : boolean; TypeChamp : string);
    procedure BasculeChampDispo(Champ : string; Indice : integer);
    procedure BasculeChampCondi(Champ : string; Indice : integer);
    procedure AjouterClick(Sender: TObject);
    procedure EnleverClick(Sender: TObject);
    procedure PrefixeClick(Sender: TObject);
    procedure ChargeListeEvent;
    procedure ParEmailClick(Sender: TObject);
    procedure NotificationClick(Sender: TObject);
    procedure MajEvenementCroise(Prefixe,Suffixe : String) ;
    function supp_suffixe(stName,Suffixe : String) : String;
    procedure AdapteLesEvenements ( Prefixe : String ) ;
    Procedure DuplicationClick( Sender : tObject);
    Procedure SetBoutonEnabled( bb : boolean);
    Procedure SetEventEnabled( numEvent : integer);
    procedure adapteValidDesti;
    end ;
Implementation
uses
  YAlertesConst,EntPgi,UtilPgi
	 ,CbpMCD
   ,CbpEnumerator
   ;

procedure TOM_YALERTES.OnNewRecord ;
var prefixe,stValeur : string;
    i,num1,num2 : integer;
begin
  Inherited ;
  SetField('YAL_DATEECHEANCE', iDate2099);
  SetField('YAL_DATEDEBUT', Date);
//  SetField ('YAL_VALIDATION','S');
  SetField ('YAL_TYPEALERTE','S');
  {SetField ('YAL_NOTIFICATION','M');}
  SetField ('YAL_DESTINATAIRES','');
  SetField('YAL_ACTIVE','X');
  SetControlChecked('YAL_ACTIVE',true);
  SetControlChecked('YAL_MEMO',true);
  SetControlVisible ('PAREMAIL',false);
  SetField('YAL_MEMO','X');
  SetField ('YAL_MODEBLOCAGE',SansBlocage);
  SetControlVisible  ('LIST_UTIL',false);
  SetControlEnabled('YAL_PREFIXE',true);
  {V_PGI.NumVersion:='8.0.0' ;}
  if not Isnumeric(copy(V_PGI.NumVersion,1,1)) or not Isnumeric(copy(V_PGI.NumVersion,3,1)) then
    SetField('YAL_VERSIONALERTE','08.000.001')
  else
    begin
    num1:=ValeurI(copy(V_PGI.NumVersion,1,1));
    num2:=ValeurI(copy(V_PGI.NumVersion,3,1));
    SetField('YAL_VERSIONALERTE',Format('%.2d',[num1])+'.'+Format('%.3d',[num2])+'.'+'001');
    end;
{ Cas création Alerte à partir d'une fiche : pré-init des champs }
  if InitChamp <> '' then
    begin
    SetField ('YAL_TYPEALERTE','S');
    prefixe:=copy(InitChamp,1,pos('_',InitChamp)-1);
    SetField ('YAL_PREFIXE',prefixe);
    SetControlEnabled ('YAL_PREFIXE',false);
    SetControlChecked('YAL_EVENEMENT1',true);
    SetField('YAL_EVENEMENT1','X');
    stValeur:=StringReplace(InitValeur,'|','-',[rfReplaceAll]);
    SetField('YAL_ALERTE',prefixe+'-'+stValeur);
    if InitLibelle <> '' then
      SetField('YAL_LIBELLE',InitLibelle)
    else
      SetField('YAL_LIBELLE',stValeur);
    TFFiche(Ecran).Pages.ActivePage:=TTabSheet(Ecran.FindComponent('BLOCNOTES'));
    end;
    for i := 1 to 3 do
      begin
      SetControlText ('YAD_CHAMP'+IntToStr(i),'');
      SetControlText ('YAD_OPER'+IntToStr(i),'');
      SetControlText ('YAD_VAL'+IntToStr(i),'');
      if i < 3 then SetControlText ('YAD_LIEN'+IntToStr(i),'');
      end;
end ;

procedure TOM_YALERTES.OnDeleteRecord ;
begin
  Inherited ;
  MajConditions(true);
  MajDestinataires(true);
end ;

procedure TOM_YALERTES.OnUpdateRecord ;
var MC : TListBox;
    OneChecked : boolean;
    i,numv : integer;
begin
  Inherited ;
  if (DS.State = dsInsert) and ( copy(GetControlText('YAL_ALERTE'),1,3) = 'CEG' ) then
    begin
    LastError:=6;
    LastErrorMsg:=TesteMessage[LastError];
    SetFocusControl('YAL_ALERTE') ;
    exit;
    end;

  if ( Trim(GetControlText ('YAL_PREFIXE')) = '' ) {and
     ( GetControlText ('YAL_PREFIXE_') = '' )} then
    begin
    LastError:=1;
    LastErrorMsg:=TesteMessage[LastError];
    SetFocusControl('YAL_PREFIXE') ;
    exit;
    end;
  MC:=TListBox(GetControl('LIST_UTIL'));
  if (TRadioButton(GetControl('UT_UTIL')).Checked = true) and
     (MC.SelCount = 0) then
    begin
    LastError:=2;
    LastErrorMsg:=TesteMessage[LastError];
    exit;
    end;
  if (TRadioButton(GetControl('UT_GRP')).Checked = true) and
     (MC.SelCount = 0) then
    begin
    LastError:=3;
    LastErrorMsg:=TesteMessage[LastError];
    exit;
    end;
  OneChecked:=false;
  for i:=1 to NbEvent do
    begin
    if Assigned(GetControl('YAL_EVENEMENT'+IntToSTr(i))) then
      begin
      if TCheckBox(GetControl('YAL_EVENEMENT'+IntToSTr(i))).Checked then
        begin
        OneChecked:=true;
        break;
        end;
      end;
    end;
  if not OneChecked then
    begin
    SetFocusControl('YAL_EVENEMENT1') ;
    LastError:=4;
    LastErrorMsg:=TesteMessage[LastError];
    exit;
    end;

  { alerte sur ouverture fiche : pas de blocage ou d'interro }
  if ( GetControlText ('YAL_MODEBLOCAGE') <> SansBlocage ) and
    ( TCheckBox(GetControl('YAL_EVENEMENT1')).Checked ) then
    begin
    LastError:=5;
    LastErrorMsg:=TesteMessage[LastError];
    SetFocusControl('YAL_MODEBLOCAGE') ;
    exit;
    end;

  { alerte sur modifs de champs : si champs infos compl, pas de blocage ou d'interro
    car les infos compl. ne se gérent pas ds les fiches principales }
  if ( GetControlText ('YAL_MODEBLOCAGE') <> SansBlocage ) and
    ( TCheckBox(GetControl('YAL_EVENEMENT5')).Checked ) then
    begin
    for i:=0 to ChampsCondi.count-1 do
      if copy(ChampsCondi.Strings[i],1,2) = 'RD' then
        begin
        LastError:=7;
        LastErrorMsg:=TesteMessage[LastError];
        SetFocusControl('YAL_MODEBLOCAGE') ;
        exit;
        end;
    end;

  { Validation simple ou J'ai lu }
  if TRadioButton(GetControl('TYPE_OK')).Checked = true then
     SetField ('YAL_VALIDATION','S')
  else
     SetField ('YAL_VALIDATION','');
  { Notification }
  {if THDBCheckBox(GetControl('NOT_MEMO')).Checked = true then
     SetField ('YAL_NOTIFICATION','M')
  else
     SetField ('YAL_NOTIFICATION','E');}

  { Alerte simple ou croisée }
  SetField ('YAL_TYPEALERTE','S');
  { Destinaires Tous, utilisateurs ou groupes }
  if TRadioButton(GetControl('UT_TOUS')).Checked = true then
     SetField ('YAL_DESTINATAIRES','')
  else
    if TRadioButton(GetControl('UT_UTIL')).Checked = true then
       SetField ('YAL_DESTINATAIRES','U')
    else
       SetField ('YAL_DESTINATAIRES','G');

  if (DS.State <> dsInsert) and IsNumeric(copy(GetField('YAL_VERSIONALERTE'),8,3)) then
    begin
    numv:=ValeurI(copy(GetField('YAL_VERSIONALERTE'),8,3));
    Inc(numv);
    SetField('YAL_VERSIONALERTE',copy(GetField('YAL_VERSIONALERTE'),1,7)+Format('%.3d',[numv]));
    end;
end ;

procedure TOM_YALERTES.OnAfterUpdateRecord ;
var Sql : string;
    T,TobAlertes : Tob;
    i : integer;
begin
  Inherited ;
  { Alertes croisées : on met à jour les tables secondaires tiers et articles }
  if ( (Assigned(GetControl('YAL_EVENEMENT6'))) and (TCheckBox(GetControl('YAL_EVENEMENT6')).Checked = true)) or
     ( (Assigned(GetControl('YAL_EVENEMENT7'))) and (TCheckBox(GetControl('YAL_EVENEMENT7')).Checked = true)) then
     MajEvenementCroise('T','_');

  if ( (Assigned(GetControl('YAL_EVENEMENT7'))) and (TCheckBox(GetControl('YAL_EVENEMENT7')).Checked = true)) then
     MajEvenementCroise('GA','__');

  { MAJ Tables utilisateurs ou groupe }
  MajDestinataires(false);
  { MAJ Tables conditions }
  MajConditions(false);
  { MAJ champs critères d'alertes }
  if (TCheckBox(GetControl('YAL_EVENEMENT4')).Checked = true) then
     MajChamps ('C');
  if (TCheckBox(GetControl('YAL_EVENEMENT5')).Checked = true) then
     MajChamps ('D');
  { si duplication on duplique les adresses mail si à dupliquer }
  if (bDuplication and
{$IFNDEF EAGLCLIENT}
    THDBCheckBox(GetControl('YAL_EMAIL')).Checked) then
{$ELSE}
    TCheckBox(GetControl('YAL_EMAIL')).Checked) then
{$ENDIF}
      if ExisteSql ('Select * from YALERTESMAIL where YAM_ALERTE="'+oldCodeAlerte+'"') then
        if not ExisteSql ('Select * from YALERTESMAIL where YAM_ALERTE="'+GetField('YAL_ALERTE')+'"') then
          begin
          Sql:='insert into YALERTESMAIL (YAM_ALERTE,YAM_MAIL,YAM_NUMERO) select "'+GetField('YAL_ALERTE')+'"'+
              ',YAM_MAIL,YAM_NUMERO from YALERTESMAIL where YAM_ALERTE="'+oldCodeAlerte+'"';
          ExecuteSql (Sql);
          end;

  { mng 27/12/2007 : on met à jour la memory tob, si modif on commence par supprimer les tobs de cette alerte }
  if VH_EntPgi.TobAlertes.detail.count = 0 then exit;
  if TFFiche(ecran).TypeAction <> taCreat then
    begin
    TobAlertes := VH_EntPgi.TobAlertes.FindFirst(['YAL_ALERTE'],
      [GetField('YAL_ALERTE')], false);
    while Assigned(TobAlertes) do
      begin
      TobAlertes.free;
      TobAlertes := VH_EntPgi.TobAlertes.FindFirst(['YAL_ALERTE'],
        [GetField('YAL_ALERTE')], false);
      end;
    end;

    sql := 'SELECT * FROM YALERTES LEFT JOIN YALERTESUTI ON YAU_ALERTE=YAL_ALERTE'+
    ' LEFT JOIN YALERTESCOND ON YAD_ALERTE=YAL_ALERTE WHERE YAL_ACTIVE="X"'+
    ' AND YAL_ALERTE = "'+GetField('YAL_ALERTE')+'"'+
    ' AND YAL_DATEECHEANCE >= "'+USDateTime(Date)+'" AND YAL_DATEDEBUT <= "'+USDateTime(Date)+
    '" AND (YAU_UTIGROUPE="'+V_PGI.USER
    +'" OR YAU_UTIGROUPE="'+V_PGI.Groupe+'" OR YAU_UTIGROUPE is NULL OR YAU_TYPE="C" OR YAU_TYPE="D")'
    ;

    TobAlertes := Tob.Create('les news alertes', nil, -1);
    TobAlertes.LoadDetailDBFromSql('YALERTES', sql);
    if TobAlertes.detail.count > 0 then
      for i:= 0 to TobAlertes.detail.count - 1 Do
        begin
        T:=Tob.Create('la fille alerte',VH_EntPgi.TobAlertes,-1);
        T.Dupliquer(TobAlertes.Detail[i], False, True);
        end;
    TobAlertes.free;
end ;

procedure TOM_YALERTES.MajEvenementCroise(Prefixe,Suffixe : String) ;
var TobCondi : tob;
    i : integer;
    PCons : TPanel;
    C : TControl ;
    stName : String;
begin
  ExecuteSQL ('delete from YALERTESCOND where YAD_ALERTE="'+GetField('YAL_ALERTE')+'" and YAD_OBJET="'+Prefixe+'"');
  if GetControlText ('YAD_CHAMP1'+Suffixe) = '' then exit;

  TobCondi:=TOB.create ('YALERTESCOND',NIL,-1);
  TobCondi.InitValeurs();
  TobCondi.SetString('YAD_ALERTE',GetField('YAL_ALERTE'));
  TobCondi.SetString('YAD_OBJET',Prefixe);

  if Suffixe = '_' then
    PCons := TPanel(getcontrol('PCONDITIERS'))
  else
    PCons := TPanel(getcontrol('PCONDIARTICLE'));

  for i:=0 to PCons.ControlCount-1 do
     begin
     C:=PCons.Controls[i] ;
     stName:=supp_suffixe(C.Name,Suffixe);
     if Copy(C.Name,5,2)='CH' then
       begin
       if THValComboBox(C).ItemIndex >= 0 then
         TobCondi.Putvalue(stName,ThValComboBox(C).Values[THValComboBox(C).ItemIndex]);
       end else
       if Copy(C.Name,5,2)='OP' then
         begin
         if THValComboBox(C).ItemIndex >= 0 then
           TobCondi.Putvalue(stName,ThValComboBox(C).Values[THValComboBox(C).ItemIndex]);
         end else
         if Copy(C.Name,5,2)='VA' then
           begin
           if (THValComboBox(C).ItemIndex = -1) then
             TobCondi.Putvalue(stName,THValComboBox(C).text)
           else
             TobCondi.Putvalue(stName,THValComboBox(C).Values[THValComboBox(C).ItemIndex]);
           end
         else
           if Copy(C.Name,5,2)='LI' then
             TobCondi.Putvalue(stName,THValComboBox(C).text);
     end;
  TobCondi.InsertOrUpdateDB;
  TobCondi.free;
  { pour le cas où l'on passe de 7 à 6 }
  if ( Prefixe = 'T' ) and (TCheckBox(GetControl('YAL_EVENEMENT6')).Checked = true) then
    ExecuteSQL ('delete from YALERTESCOND where YAD_ALERTE="'+GetField('YAL_ALERTE')+'" and YAD_OBJET="GA"');
end;

procedure TOM_YALERTES.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_YALERTES.OnLoadRecord ;
begin
  Inherited ;
  
  if GetField('YAL_PREDEFINI') = 'CEG' then
    begin
    FicheReadOnly(Ecran,true);
    stAction:='CONSULTATION';
    end
  else
    if ( stAction = 'CONSULTATION' ) and (GetField('YAL_PREDEFINI') <> 'CEG' ) then
      SetControlEnabled('BDUPLICATION',false);
  if ( stAction = 'CONSULTATION' ) then
    SetBoutonEnabled(false);
  if (DS.State = dsInsert) and ( not bDuplication) then
    SetField ('YAL_VALIDATION','S');
  adapteValidDesti;
  ChargeListeConditions(GetField('YAL_PREFIXE'),'');
  { Destinaires Tous, utilisateurs ou groupes }
  if (GetField ('YAL_EVENEMENT4') = 'X') or (GetField ('YAL_EVENEMENT5') = 'X') then
    SetControlVisible  ('SELECTIONCHAMP',true);
  if ( Assigned(GetControl('YAL_EVENEMENT6'))) and (GetField ('YAL_EVENEMENT6') = 'X') then
    begin
    SetControlVisible  ('PCONDITIERS',true);
    ChargeListeConditions('T','_');
    end;
  if ( Assigned(GetControl('YAL_EVENEMENT7'))) and (GetField ('YAL_EVENEMENT7') = 'X') then
    begin
    SetControlVisible  ('PCONDITIERS',true);
    SetControlVisible  ('PCONDIARTICLE',true);
    ChargeListeConditions('T','_');
    ChargeListeConditions('GA','__');
    end;

  if ( not AfterInserting ) then
    begin
    SetControlEnabled('YAL_PREFIXE',false);
    //SetControlEnabled('YAL_PREFIXE_',false);
    PrefixeClick(Ecran);
    //if GetField ('YAL_TYPEALERTE') <> 'S' then SetControlText ('YAL_PREFIXE_',GetField('YAL_PREFIXE'));
    end;

  stArgument:=false;

end ;

procedure TOM_YALERTES.ChargeListeConditions (Prefixe,Suffixe : String);
var ListeChamp,ListeLibelle: HTStrings;
    TobCondi : tob;
    QQ : TQuery;
    C : TControl ;
    CBValeur : THValComboBox ;
    ChampLibre,CValeur,stName,stDatatype,stChamps,Champ,stValeurs,Valeur : string;
    i ,index: integer;
    PCons : TPanel;
begin
  { partie conditions }
  ListeChamp := HTStringList.Create;
  ListeLibelle := HTStringList.Create;
  ListeChamp.add (''); ListeLibelle.add (TraduireMemoire('<<Aucun>>'));
  chargeComboChamps (Prefixe,'L','',ListeChamp,ListeLibelle);

  if Suffixe='' then
    begin
    CBCHAMP1.Items.Assign(ListeLibelle);
    CBCHAMP1.Values.Assign(ListeChamp);
    CBCHAMP2.Items.Assign(ListeLibelle);
    CBCHAMP2.Values.Assign(ListeChamp);
    CBCHAMP3.Items.Assign(ListeLibelle);
    CBCHAMP3.Values.Assign(ListeChamp);
    end
  else
    if Suffixe='_' then
      begin
      CBCHAMP1_.Items.Assign(ListeLibelle);
      CBCHAMP1_.Values.Assign(ListeChamp);
      CBCHAMP2_.Items.Assign(ListeLibelle);
      CBCHAMP2_.Values.Assign(ListeChamp);
      CBCHAMP3_.Items.Assign(ListeLibelle);
      CBCHAMP3_.Values.Assign(ListeChamp);
      end
    else
      begin
      CBCHAMP1__.Items.Assign(ListeLibelle);
      CBCHAMP1__.Values.Assign(ListeChamp);
      CBCHAMP2__.Items.Assign(ListeLibelle);
      CBCHAMP2__.Values.Assign(ListeChamp);
      CBCHAMP3__.Items.Assign(ListeLibelle);
      CBCHAMP3__.Values.Assign(ListeChamp);
      end;

  ListeChamp.free;ListeLibelle.Free;
  TobCondi:=TOB.create ('YALERTESCOND',NIL,-1);
  QQ := OpenSQL('Select * from YALERTESCOND where YAD_ALERTE="'+GetControlText('YAL_ALERTE')+'" AND '+
     'YAD_OBJET="'+Prefixe+'" ',true);
  if Not QQ.EOF then
  begin
    TobCondi.SelectDB('',QQ);
      if Suffixe = '' then
        PCons := TPanel(getcontrol('PCONDI'))
      else
        if Suffixe = '_' then
          PCons := TPanel(getcontrol('PCONDITIERS'))
        else
          PCons := TPanel(getcontrol('PCONDIARTICLE'));

      for i:=0 to PCons.ControlCount-1 do
      begin
         C:=PCons.Controls[i] ;
         if not Assigned(C) then continue;
         { pour le tob.getvalue, on enleve le suffixe du nom du champ }
         stName:=supp_suffixe(C.Name,Suffixe);
         { controle écran }
         CValeur := FindEtReplace (C.Name,'CHAMP','VAL',false);
         CBValeur := THValcomboBox(GetControl(CValeur));
         { champ de la table }
         CValeur := FindEtReplace (stName,'CHAMP','VAL',false);
         if Copy(stName,5,2)='CH' then
         begin
            ChampLibre := TobCondi.GetValue(stName);
            THValcomboBox(GetControl(C.Name)).value := ChampLibre;
            CBValeur.datatype := '';  CBValeur.text := '';
            CBValeur.values.clear;    CBValeur.items.clear;
            if (pos ('LIBRETIERS',ChampLibre)>0) or (ChampLibre='T_REPRESENTANT') or (ChampToType( ChampLibre) = 'COMBO') then
            begin
              stDatatype := AffecteDatatype(ChampLibre, '');
              CBValeur.datatype:=stDatatype;
              index := CBValeur.Values.IndexOf(TobCondi.GetValue(CValeur));
              if (index>-1) then
              begin
                 CBValeur.value := TobCondi.GetValue(CValeur);
                 CBValeur.text := CBValeur.Items[index];
              end else
                 CBValeur.text := TobCondi.GetValue(CValeur);
            end else CBValeur.text := TobCondi.GetValue(CValeur);
         end else
         if ((Copy(stName,5,2)='OP') or (Copy(stName,5,2)='LI')) then
         begin
            THValcomboBox(GetControl(C.Name)).value := TobCondi.GetValue(stName);
         end;
      end;
  end;
  ferme (QQ);
  { init première ligne si création à partir d'une fiche }
  if (AfterInserting) and (InitChamp <> '' ) then
    begin
    stChamps:=InitChamp;
    stValeurs:=InitValeur;
    i:=0;
    Repeat
      Valeur:=ReadToKenPipe(stValeurs,'|');
      Champ:=ReadToKenPipe(stChamps,'|');
      if Valeur <> '' then
        begin
        inc(i);
        THValcomboBox(GetControl('YAD_CHAMP'+IntToStr(i))).value:=Champ;
        THValcomboBox(GetControl('YAD_OPER'+IntToStr(i))).value:='=';
        THValcomboBox(GetControl('YAD_VAL'+IntToStr(i))).text:=Valeur;
        if i > 1 then
          THValcomboBox(GetControl('YAD_LIEN'+IntToStr(i-1))).value:='Et';
        end;
    until Valeur='';

    end;
  TobCondi.free;
end;

procedure TOM_YALERTES.AdapteLesEvenements ( Prefixe : String ) ;
var i : integer;
    bActiv : boolean;
begin
  { pour l'instant les event 6 et 7 ne valent que pour les pièces }
  if Prefixe <> 'GP' then bActiv:=false else bActiv:=true;
  for i:=6 to 7 do
    if Assigned(GetControl('YAL_EVENEMENT'+IntToSTr(i))) then
      TCheckBox(GetControl('YAL_EVENEMENT'+IntToSTr(i))).Enabled := bActiv;
end;

procedure TOM_YALERTES.OnArgument ( S: String ) ;
var stValeurs,Critere,ChampMul,ValMul,stChp1,stChp2 : String;
    x,i,j : integer;
begin
  Inherited ;
  bMultiLine:=V_PGI.MultiLineOnglet;
  if V_PGI.MultiLineOnglet = true then
    V_PGI.MultiLineOnglet:=false;
{$IFNDEF GPAO}
  { Alertes - Tables gérées seulement en GP : Wxx et GPO, GEM, GDE, YTS }
  if Assigned(GetControl('YAL_PREFIXE')) then
    THValcomboBox(GetControl('YAL_PREFIXE')).Plus := ' AND ' + WhereAlertesNotGP;
{$ENDIF GPAO}

  stAction :=GetArgumentValue(S, 'ACTION');
  stArgument:=true;
  bDuplication  := false;
  oldCodeAlerte:='';
  ChargeListeEvent;
  if pos ('LOG',S) > 0 then
    Ecran.Caption:='Dernière version de l''alerte';
  TRadioButton(GetControl('UT_UTIL')).OnClick:=DestiClick;
  TRadioButton(GetControl('UT_GRP')).OnClick:=DestiClick;
  TRadioButton(GetControl('UT_TOUS')).OnClick:=DestiClick;
  TRadioButton(GetControl('TYPE_OK')).OnClick:=TypeClick;
  TRadioButton(GetControl('TYPE_JLOK')).OnClick:=TypeClick;
  if Assigned(GetControl('YAL_EVENEMENT1')) then
    TCheckBox(GetControl('YAL_EVENEMENT1')).OnClick:=Evenement1Click;
  if Assigned(GetControl('YAL_EVENEMENT2')) then
    TCheckBox(GetControl('YAL_EVENEMENT2')).OnClick:=Evenement2Click;
  if Assigned(GetControl('YAL_EVENEMENT3')) then
    TCheckBox(GetControl('YAL_EVENEMENT3')).OnClick:=Evenement3Click;
  if Assigned(GetControl('YAL_EVENEMENT4')) then
    TCheckBox(GetControl('YAL_EVENEMENT4')).OnClick:=Evenement4Click;
  if Assigned(GetControl('YAL_EVENEMENT5')) then
    TCheckBox(GetControl('YAL_EVENEMENT5')).OnClick:=Evenement5Click;
  if Assigned(GetControl('YAL_EVENEMENT6')) then
    TCheckBox(GetControl('YAL_EVENEMENT6')).OnClick:=Evenement6Click;
  if Assigned(GetControl('YAL_EVENEMENT7')) then
    TCheckBox(GetControl('YAL_EVENEMENT7')).OnClick:=Evenement7Click;
  if Assigned(GetControl('YAL_EVENEMENT8')) then
    TCheckBox(GetControl('YAL_EVENEMENT8')).OnClick:=Evenement8Click;
  if Assigned(GetControl('LIST_UTIL')) then
    TListBox(GetControl('LIST_UTIL')).OnClick := ZoneClick;

  TToolBarButton97(GetControl('BAJOUTER')).OnClick:=AjouterClick;
  TToolBarButton97(GetControl('BENLEVER')).OnClick:=EnleverClick;
  //TToolBarButton97(GetControl('Bdefaire')).OnClick:=bDefaireClick;
  THValcomboBox(GetControl('YAL_PREFIXE')).OnClick:=PrefixeClick;
  TBitBtn(GetControl('PAREMAIL')).OnClick:=ParEmailClick;
  {THDBCheckBox(GetControl('NOT_MEMO')).OnClick:=NotificationClick;}
{$IFNDEF EAGLCLIENT}
  THDBCheckBox(GetControl('YAL_EMAIL')).OnClick:=NotificationClick;
{$ELSE}
  TCheckBox(GetControl('YAL_EMAIL')).OnClick:=NotificationClick;
{$ENDIF}
  { duplication }
  if Assigned(GetControl('BDUPLICATION')) then
    TToolBarButton97(GetControl('BDUPLICATION')).OnClick := DuplicationClick;
  { partie conditions }
  CBCHAMP1 := THValcomboBox(GetControl('YAD_CHAMP1'));
  CBCHAMP1.OnClick := ALCHAMP1Click;
  CBCHAMP2 := THValcomboBox(GetControl('YAD_CHAMP2'));
  CBCHAMP2.OnClick := ALCHAMP2Click;
  CBCHAMP3 := THValcomboBox(GetControl('YAD_CHAMP3'));
  CBCHAMP3.OnClick := ALCHAMP3Click;
  CBCHAMP1_ := THValcomboBox(GetControl('YAD_CHAMP1_'));
  CBCHAMP1_.OnClick := ALCHAMP1_Click;
  CBCHAMP2_ := THValcomboBox(GetControl('YAD_CHAMP2_'));
  CBCHAMP2_.OnClick := ALCHAMP2_Click;
  CBCHAMP3_ := THValcomboBox(GetControl('YAD_CHAMP3_'));
  CBCHAMP3_.OnClick := ALCHAMP3_Click;
  CBCHAMP1__ := THValcomboBox(GetControl('YAD_CHAMP1__'));
  CBCHAMP1__.OnClick := ALCHAMP1__Click;
  CBCHAMP2__ := THValcomboBox(GetControl('YAD_CHAMP2__'));
  CBCHAMP2__.OnClick := ALCHAMP2__Click;
  CBCHAMP3__ := THValcomboBox(GetControl('YAD_CHAMP3__'));
  CBCHAMP3__.OnClick := ALCHAMP3__Click;

  stChp1:=stInitChp1;
  stChp2:=stInitChp2;

  for j:= 1 to 3 Do
    for i :=1 to 3 Do
      begin
      stChp1:=stChp1+IntToStr(i);
      if j = 2 then stChp1:=stChp1+'_' else if j = 3 then stChp1:=stChp1+'__' ;
      if Assigned(GetControl(stChp1)) then
        THValcomboBox(GetControl(stChp1)).OnExit := ZoneClick;
      stChp2:=stChp2+IntToStr(i);
      if j = 2 then stChp2:=stChp2+'_' else if j = 3 then stChp2:=stChp2+'__' ;
      if Assigned(GetControl(stChp2)) then
        THValcomboBox(GetControl(stChp2)).OnExit := ZoneClick;
      stChp1:=stInitChp1;
      stChp2:=stInitChp2;
      end;

  ChampsDispo := HTStringList.Create;
  ChampsCondi := HTStringList.Create;
  { Cas création Alerte à partir d'une fiche : pré-init des champs }
  InitChamp:=''; InitValeur:='';InitLibelle:='';
  stValeurs:=S;
  Repeat
    Critere:=ReadToKenSt(stValeurs);
    if Critere <> '' then
      begin
        x:=pos('=',Critere);
        if x<>0 then
        begin
          ChampMul:=copy(Critere,1,x-1);
          ValMul:=copy(Critere,x+1,length(Critere));
          if ChampMul='CHAMP' then InitChamp:=ValMul;
          if ChampMul='VALEUR' then InitValeur:=ValMul;
          if ChampMul='LIBELLE' then InitLibelle:=ValMul;
        end;
      end;
  until Critere='';
end ;

procedure TOM_YALERTES.OnClose ;
begin
  Inherited ;
  if bMultiLine then V_PGI.MultiLineOnglet:=true;
  if Assigned(TobDesti) then FreeAndNil(TobDesti);
  if Assigned(ChampsDispo) then FreeAndNil(ChampsDispo);
  if Assigned(ChampsCondi) then FreeAndNil(ChampsCondi);
end ;

procedure TOM_YALERTES.OnCancelRecord ;
begin
  Inherited ;
  if bDuplication then
    begin
    bDuplication := false;
    oldCodeAlerte:='';
    SetControlEnabled('BDUPLICATION',true);
    end;
end ;

procedure TOM_YALERTES.ChgtDestinataires ;
begin
  if not stArgument then ForceUpdate;
  if TRadioButton(GetControl('UT_TOUS')).Checked = true then
     SetControlVisible('LIST_UTIL',false)
  else
     ChargeListUser();
end;

procedure TOM_YALERTES.ChargeListUser ();
var ii,indice,iElement,Pos1,i,j : integer;
    stCombos,NomChamp : String ;
    Valeur : hString;
    StString : hstring;
    ListeGroup : HTStringList;
    StChamp : array[1..3] of hstring;
    MC : THListBox;
begin
  SetControlVisible  ('LIST_UTIL',true);
  MC:=THListBox(GetControl('LIST_UTIL'));
  MC.Items.Clear;

  if (TRadioButton(GetControl('UT_UTIL')).Checked = true ) then
     begin
     //NomChamp:='YAU_UTILISATEUR';
     stCombos:=stComboUtilisateur;
     end
  else
     begin
     //NomChamp:='YAU_GROUPE';
     stCombos:=stComboGroupe;
     end;
  NomChamp:='YAU_UTIGROUPE';
  ii:=TTToNum(stCombos) ;
  if ii<0 then exit;

  if (V_PGI.DECombos[ii].Valeurs = Nil) then RemplirListe (stCombos,'');
  if (V_PGI.DECombos[ii].Valeurs = Nil) then exit;

  ListeGroup := hTStringList.Create;
  for i:=0 to Pred(V_PGI.DECombos[ii].Valeurs.count) do ListeGroup.Add (V_PGI.DECombos[ii].Valeurs.Strings[i]);

  { je n'utilise pas directement cette liste (viol acces), je la duplique }
  iElement := 0;
  While iElement < ListeGroup.count do
     begin
     StString := ListeGroup.Strings[iElement];
     for indice := 0 to 2 do
         begin
         Pos1 := Pos (#9,StString);
         if Pos1 > 1 then
            begin
            StChamp[indice+1] := Copy(StString,1,Pos1-1);Delete(StString,1,Pos1);
            end
         else
            begin
            StChamp[indice+1] := '';
            end;
         end;
     Valeur:=StChamp[1]+' '+StChamp[2];
     if Valeur <> ' ' then
        MC.Items.Add(Valeur);
     inc (iElement)
     end;

  { select des lignes en tob }
  if ( TobDesti<> Nil ) and ( TobDesti.detail.count > 0 ) and (ListeGroup.count > 0 ) then
    for i:=0 to TobDesti.detail.count-1 do
      for j:=0 to ListeGroup.count-1 do
        begin
        if Copy(MC.Items[j],1,Pos(' ',MC.Items[j])-1)=TobDesti.detail[i].GetValue(NomChamp) then
           begin
           MC.Selected[j]:=True;
           break;
           end;
        end;
  ListeGroup.Free;
end;

procedure TOM_YALERTES.AlimTobDestinataires();
var Q : TQuery;
begin
  if Assigned(TobDesti) then TobDesti.free;
  TobDesti:=TOB.create ('les destinataires',NIL,-1);
  if ( ( not AfterInserting ) or (bDuplication) ) then
    begin
    Q := OpenSql ('Select * from YALERTESUTI where YAU_ALERTE="'+GetControlText('YAL_ALERTE')+'"',TRUE);
    TobDesti.LoadDetailDB ('YALERTESUTI', '', '', Q, False);
    ferme(Q) ;
    end;
end;

procedure TOM_YALERTES.MajDestinataires(delete : boolean);
var i : integer;
    stUtil : String;
    MC : TListBox;
begin
    ExecuteSQL ('delete from YALERTESUTI where YAU_ALERTE="'+GetField('YAL_ALERTE')+'"');
    if delete then exit;
    MC:=TListBox(GetControl('LIST_UTIL'));
    if (TRadioButton(GetControl('UT_TOUS')).Checked = false) and (MC.Items.count <> 0 ) then
      begin
      for i:=0 to MC.Items.count-1 do
        if MC.Selected[i]=True then
          begin
          stUtil:=Copy(MC.Items[i],1,Pos(' ',MC.Items[i])-1);
          ExecuteSQL ('INSERT INTO YALERTESUTI (YAU_ALERTE,YAU_TYPE,YAU_UTIGROUPE,'
          +'YAU_CHAMP) values ("'+GetField('YAL_ALERTE')+'","'+GetField('YAL_DESTINATAIRES')+'",'
          +'"'+stUtil+'","")');
          end;
      end;
end;

procedure TOM_YALERTES.MajConditions(delete : boolean);
var i : integer;
    tobcondi : tob;
    PCons : TPanel;
    C : TControl ;
begin
    ExecuteSQL ('delete from YALERTESCOND where YAD_ALERTE="'+GetField('YAL_ALERTE')+'" AND '+
     'YAD_OBJET="'+GetField('YAL_PREFIXE')+'" ');
    if delete then
      begin
      { si croisée 6 ou 7, supp objet 'T' }
      if (TCheckBox(GetControl('YAL_EVENEMENT6')).Checked = true) or
         (TCheckBox(GetControl('YAL_EVENEMENT7')).Checked = true) then
        ExecuteSQL ('delete from YALERTESCOND where YAD_ALERTE="'+GetField('YAL_ALERTE')+'" AND '+
         'YAD_OBJET="T" ');
      { si croisée 7, supp objet 'GA' }
      if (TCheckBox(GetControl('YAL_EVENEMENT7')).Checked = true) then
        ExecuteSQL ('delete from YALERTESCOND where YAD_ALERTE="'+GetField('YAL_ALERTE')+'" AND '+
         'YAD_OBJET="GA" ');
      exit;
      end;
    if NbEvent >= 4 then
      if (TCheckBox(GetControl('YAL_EVENEMENT4')).Checked = true) or
         (TCheckBox(GetControl('YAL_EVENEMENT5')).Checked = true) then exit;
    TobCondi:=TOB.create ('YALERTESCOND',NIL,-1);
    if CBCHAMP1.Value <> '' then
      begin
      TobCondi.Putvalue('YAD_ALERTE',GetField('YAL_ALERTE'));
      TobCondi.Putvalue('YAD_OBJET',GetField('YAL_PREFIXE'));
      PCons := TPanel(getcontrol('PCONDI'));
      for i:=0 to PCons.ControlCount-1 do
         begin
         C:=PCons.Controls[i] ;
         if Copy(C.Name,5,2)='CH' then
           begin
           if THValComboBox(C).ItemIndex >= 0 then
             TobCondi.Putvalue(C.Name,ThValComboBox(C).Values[THValComboBox(C).ItemIndex]);
           end else
           if Copy(C.Name,5,2)='OP' then
             begin
             if THValComboBox(C).ItemIndex >= 0 then
               TobCondi.Putvalue(C.Name,ThValComboBox(C).Values[THValComboBox(C).ItemIndex]);
             end else
             if Copy(C.Name,5,2)='VA' then
               begin
               if (THValComboBox(C).ItemIndex = -1) then
                 TobCondi.Putvalue(C.Name,THValComboBox(C).text)
               else
                 TobCondi.Putvalue(C.Name,THValComboBox(C).Values[THValComboBox(C).ItemIndex]);
               end else
               if Copy(C.Name,5,2)='LI' then
                 if THValComboBox(C).ItemIndex >= 0 then
                   TobCondi.Putvalue(C.Name,THValComboBox(C).text);
             end;
      TobCondi.InsertOrUpdateDB;
      end;
    TobCondi.free;
end;

procedure TOM_YALERTES.MajChamps(TypeChamp : String);
var i : integer;
begin
  ExecuteSQL ('delete from YALERTESUTI where YAU_ALERTE="'+GetField('YAL_ALERTE')+'"'+
     ' AND YAU_TYPE="'+TypeChamp+'"');
  for i:=0 to ChampsCondi.count-1 do
    begin
    ExecuteSQL ('INSERT INTO YALERTESUTI (YAU_ALERTE,YAU_TYPE,YAU_UTIGROUPE,'
    +'YAU_CHAMP) values ("'+GetField('YAL_ALERTE')+'","'+TypeChamp+'","",'
    +'"'+ChampsCondi.Strings[i]+'")');
    end;
end;

procedure TOM_YALERTES.DestiClick(Sender: TObject);
begin
  ChgtDestinataires();
end;

procedure TOM_YALERTES.TypeClick(Sender: TObject);
begin
  if not stArgument then ForceUpdate;
end;

procedure TOM_YALERTES.NotificationClick(Sender: TObject);
begin
{$IFNDEF EAGLCLIENT}
   if THDBCheckBox(GetControl('YAL_EMAIL')).Checked = false then
{$ELSE}
   if TCheckBox(GetControl('YAL_EMAIL')).Checked = false then
{$ENDIF}
    begin
    SetControlVisible ('PAREMAIL',false);
    ExecuteSQL ('DELETE FROM YALERTESMAIL WHERE YAM_ALERTE="'+GetField('YAL_ALERTE')+'"');
    end
  else
     SetControlVisible ('PAREMAIL',true);

end;

procedure TOM_YALERTES.Evenement1Click(Sender: TObject);
begin
  if not stArgument then ForceUpdate;
  if TCheckBox(GetControl('YAL_EVENEMENT1')).Checked = true then
    SetEventEnabled(1);
end;

procedure TOM_YALERTES.Evenement2Click(Sender: TObject);
begin
  if not stArgument then ForceUpdate;
  if TCheckBox(GetControl('YAL_EVENEMENT2')).Checked = true then
    SetEventEnabled(2);
end;
procedure TOM_YALERTES.Evenement3Click(Sender: TObject);
begin
  if not stArgument then ForceUpdate;
  if TCheckBox(GetControl('YAL_EVENEMENT3')).Checked = true then
    SetEventEnabled(3);
end;

procedure TOM_YALERTES.Evenement4Click(Sender: TObject);
begin
  if NbEvent=0 then exit;
  if not stArgument then ForceUpdate;
  if TCheckBox(GetControl('YAL_EVENEMENT4')).Checked = true then
     begin
     AlimListesChamps(false,'C');
     SetEventEnabled(4);
     end
  else
     AdapteLesEvenements(GetField('YAL_PREFIXE'));
end;

procedure TOM_YALERTES.Evenement5Click(Sender: TObject);
begin
  if NbEvent=0 then exit;
  if not stArgument then ForceUpdate;
  if TCheckBox(GetControl('YAL_EVENEMENT5')).Checked = true then
     begin
     AlimListesChamps(false,'D');
     SetEventEnabled(5);
     end
  else
     begin
     SetControlVisible  ('YAL_NBJOURS',false);
     AdapteLesEvenements(GetField('YAL_PREFIXE'));
     end;
end;

procedure TOM_YALERTES.Evenement6Click(Sender: TObject);
begin
  if not stArgument then ForceUpdate;

  if TCheckBox(GetControl('YAL_EVENEMENT6')).Checked = true then
     begin
     SetControlVisible  ('PCONDITIERS',true);
     ChargeListeConditions('T','_');
     SetEventEnabled(6);
     end
  else
    if not TCheckBox(GetControl('YAL_EVENEMENT7')).Checked then
      SetControlVisible  ('PCONDITIERS',false);
end;

procedure TOM_YALERTES.Evenement7Click(Sender: TObject);
begin
  if NbEvent=0 then exit;
  if not stArgument then ForceUpdate;

  if TCheckBox(GetControl('YAL_EVENEMENT7')).Checked = true then
     begin
     SetEventEnabled(7);
     SetControlVisible  ('PCONDITIERS',true);
     SetControlVisible  ('PCONDIARTICLE',true);
     ChargeListeConditions('T','_');
     ChargeListeConditions('GA','__');
     end
  else
     begin
     if not TCheckBox(GetControl('YAL_EVENEMENT6')).Checked then
        SetControlVisible  ('PCONDITIERS',false);
     SetControlVisible  ('PCONDIARTICLE',false);
     end;
end;

procedure TOM_YALERTES.Evenement8Click(Sender: TObject);
begin
  if not stArgument then ForceUpdate;
  SetEventEnabled(8);
end;

procedure TOM_YALERTES.PrefixeClick(Sender: TObject);
var stPrefixe : String;
  //Titre : String;
begin
  if NbEvent < 4 then exit;
  if (stAction <> 'CONSULTATION')  then
    begin
    SetControlEnabled('YAL_EVENEMENT4',true);
    SetControlEnabled('YAL_EVENEMENT5',true);
    end;
  stPrefixe:=GetControlText('YAL_PREFIXE');
  if Pos(stPrefixe,Prefixe45) < 0 then
    begin
    SetControlEnabled('YAL_EVENEMENT4',false);
    SetControlEnabled('YAL_EVENEMENT5',false);
    end;

  ChargeListeConditions(stPrefixe,'');
  if Assigned(GetControl('LTTABLE')) then
    begin
    SetControlCaption('LTTABLE',TraduireMemoire('Sélection ')+ TraduireMemoire (AnsiLowerCase_(RechDom('WPREFIXE',stPrefixe,false))));
    //Titre :='Sélection '+AnsiLowerCase(RechDom('WPREFIXE',stPrefixe,false));
    //SetControlCaption('LTTABLE',TraduireMemoire (Titre));
    end;
  AdapteLesEvenements(stPrefixe);
end;

procedure TOM_YALERTES.ParEmailClick(Sender: TObject);
begin
  if not stArgument then ForceUpdate;
  AGLLanceFiche('Y','YALERTES_MAIL','','',getField('YAL_ALERTE'));
end;

procedure TOM_YALERTES.AlimListesChamps(tous : boolean; TypeChamp : string);
var Q : TQuery;
    TobCondi : tob;
    i : integer;
begin
  ChampsDispo.clear;
  ChampsCondi.clear;

  ChpsDispo:=THListBox(GetControl('CHPSDISPO'));
  LibellesDispo:=ChpsDispo.Items;
  LibellesDispo.clear;
  ChpsDispo.Sorted:=True;

  //ChampsDispo.Sorted:=True;
  ChampsDispo.Duplicates:=dupIgnore;

  ChpsCondi:=THListBox(GetControl('CHPSCONDI'));
  LibellesCondi := ChpsCondi.Items;
  LibellesCondi.clear;
  ChpsCondi.Sorted:=True;

  //ChampsCondi.Sorted:=True;
  ChampsCondi.Duplicates:=dupIgnore;

  chargeComboChamps (GetField ('YAL_PREFIXE'),'L',TypeChamp,ChampsDispo,LibellesDispo);
  if tous then exit;
  { remplissage des champs condition }
  TobCondi:=TOB.create ('les champs alertes',NIL,-1);
  Q := OpenSql ('Select * from YALERTESUTI where YAU_ALERTE="'+GetField('YAL_ALERTE')+'"'+
     ' AND YAU_TYPE="'+TypeChamp+'"',TRUE);
  TobCondi.LoadDetailDB ('YALERTESUTI', '', '', Q, False);
  ferme(Q) ;
  if TobCondi.detail.count > 0 then
    For i:=0 to TobCondi.detail.count-1 Do
      BasculeChampDispo(TobCondi.detail[i].GetValue('YAU_CHAMP'),-1);
  TobCondi.free;
end;

procedure TOM_YALERTES.BasculeChampDispo(Champ : string; Indice : integer);
var Posit : integer;
begin
  // function Find(const S: string; var Index: Integer): Boolean; virtual;
  //if not ChampsDispo.find(Champ,Posit) then exit;
  if Indice = -1 then
    begin
    for Posit:=0 to ChampsDispo.Count-1 do
       if ChampsDispo.Strings[Posit]=Champ then break;
    if Posit >= ChampsDispo.Count then exit;
    end
  else Posit := Indice;

  ChampsCondi.add (Champ);
  LibellesCondi.add (LibellesDispo.Strings[Posit]);
  ChampsDispo.Delete(Posit);
  LibellesDispo.Delete(Posit);
end;

procedure TOM_YALERTES.BasculeChampCondi(Champ : string; Indice : integer);
var Posit : integer;
begin
  // function Find(const S: string; var Index: Integer): Boolean; virtual;
  { bien tenté mais j'ai enlevé le sort ...if not ChampsCondi.find(Champ,Posit) then exit;}
  if Indice = -1 then
    begin
    for Posit:=0 to ChampsCondi.Count-1 do
       if ChampsCondi.Strings[Posit]=Champ then break;
    end
  else Posit := Indice;
  if Posit >= ChampsCondi.Count then exit;
  ChampsDispo.add (Champ);
  LibellesDispo.add (LibellesCondi.Strings[Posit]);
  ChampsCondi.Delete(Posit);
  LibellesCondi.Delete(Posit);
end;

procedure TOM_YALERTES.AjouterClick(Sender: TObject);
begin
  if (ChpsDispo.ItemIndex = -1 ) or ( (ChpsDispo.ItemIndex >= 0) and
      ( ChpsDispo.Selected[ChpsDispo.ItemIndex] = False) ) then
    begin
    PGIBox('Vous devez sélectionner un champ disponible','Ajouter un champ');
    exit;
    end;
  if not AfterInserting then ForceUpdate;
  BasculeChampDispo(ChampsDispo.Strings[ChpsDispo.ItemIndex],ChpsDispo.ItemIndex);
end;

procedure TOM_YALERTES.EnleverClick(Sender: TObject);
begin
if (ChpsCondi.ItemIndex = -1  ) or (ChpsCondi.SelCount = 0) or
   ( (ChpsCondi.ItemIndex >= 0) and ( ChpsCondi.Selected[ChpsCondi.ItemIndex] = False) ) then
   begin
   PGIBox('Vous devez sélectionner un champ condition','Enlever un champ');
   exit;
   end;
  if not AfterInserting then ForceUpdate;
  BasculeChampCondi(ChampsCondi.Strings[ChpsCondi.ItemIndex],ChpsCondi.ItemIndex)
end;

{procedure TOM_YALERTES.bDefaireClick(Sender: TObject);
begin
  Inherited ;
  if bDuplication then
    begin
    bDuplication := false;
    oldCodeAlerte:='';
    end;
end;}

procedure TOM_YALERTES.ALCHAMP1Click(Sender: TObject);
var ChampLibre :  string;
begin
   if not stArgument then ForceUpdate;
   ChampLibre := THValcomboBox(GetControl('YAD_CHAMP1')).value;
   AffecteDatatype(ChampLibre,'YAD_VAL1');
end;

procedure TOM_YALERTES.ALCHAMP2Click(Sender: TObject);
var ChampLibre : string;
begin
   if not stArgument then ForceUpdate;
   ChampLibre := THValcomboBox(GetControl('YAD_CHAMP2')).value;
   AffecteDatatype(ChampLibre,'YAD_VAL2');
end;

procedure TOM_YALERTES.ALCHAMP3Click(Sender: TObject);
var ChampLibre : string;
begin
   if not stArgument then ForceUpdate;
   ChampLibre := THValcomboBox(GetControl('YAD_CHAMP3')).value;
   AffecteDatatype(ChampLibre,'YAD_VAL3');
end;

procedure TOM_YALERTES.ALCHAMP1_Click(Sender: TObject);
var ChampLibre :  string;
begin
   if not stArgument then ForceUpdate;
   ChampLibre := THValcomboBox(GetControl('YAD_CHAMP1_')).value;
   AffecteDatatype(ChampLibre,'YAD_VAL1_');
end;

procedure TOM_YALERTES.ALCHAMP2_Click(Sender: TObject);
var ChampLibre : string;
begin
   if not stArgument then ForceUpdate;
   ChampLibre := THValcomboBox(GetControl('YAD_CHAMP2_')).value;
   AffecteDatatype(ChampLibre,'YAD_VAL2_');
end;

procedure TOM_YALERTES.ALCHAMP3_Click(Sender: TObject);
var ChampLibre : string;
begin
   if not stArgument then ForceUpdate;
   ChampLibre := THValcomboBox(GetControl('YAD_CHAMP3_')).value;
   AffecteDatatype(ChampLibre,'YAD_VAL3_');
end;

procedure TOM_YALERTES.ALCHAMP1__Click(Sender: TObject);
var ChampLibre :  string;
begin
   if not stArgument then ForceUpdate;
   ChampLibre := THValcomboBox(GetControl('YAD_CHAMP1__')).value;
   AffecteDatatype(ChampLibre,'YAD_VAL1__');
end;

procedure TOM_YALERTES.ALCHAMP2__Click(Sender: TObject);
var ChampLibre : string;
begin
   if not stArgument then ForceUpdate;
   ChampLibre := THValcomboBox(GetControl('YAD_CHAMP2__')).value;
   AffecteDatatype(ChampLibre,'YAD_VAL2__');
end;

procedure TOM_YALERTES.ALCHAMP3__Click(Sender: TObject);
var ChampLibre : string;
begin
   if not stArgument then ForceUpdate;
   ChampLibre := THValcomboBox(GetControl('YAD_CHAMP3__')).value;
   AffecteDatatype(ChampLibre,'YAD_VAL3__');
end;

procedure TOM_YALERTES.ZoneClick(Sender: TObject);
begin
   if not stArgument then ForceUpdate;
end;

function TOM_YALERTES.AffecteDatatype(CHAMPTABLE,CHAMPCONF: string):string;
var datatype :string;
begin
   result := '';
   if  (CHAMPTABLE = '') then datatype := ''
   else if copy(CHAMPTABLE,1,19)='YTC_TABLELIBRETIERS' then    //YTC_TABLELIBRETIERS0 -> GCLIBRETIERS0
      datatype := FindEtReplace(CHAMPTABLE,'YTC_TABLE','GC',false)
   else if copy(CHAMPTABLE,1,15)='RPR_RPRLIBTABLE' then   //RPR_RPRLIBTABLE0 -> RT RPRLIBTABLE0
      datatype := FindEtReplace(CHAMPTABLE,'RPR_','RT',false)
   else if (CHAMPTABLE='T_REPRESENTANT') then
      datatype := 'GCCOMMERCIAL'
   else
      datatype := Get_Join(CHAMPTABLE );
   if (CHAMPCONF <> '') then
   begin
     if (THValcomboBox(GetControl(CHAMPCONF)).datatype <> datatype) then
     begin
       THValcomboBox(GetControl(CHAMPCONF)).values.clear;
       THValcomboBox(GetControl(CHAMPCONF)).items.clear;
       THValcomboBox(GetControl(CHAMPCONF)).text := '';
       THValcomboBox(GetControl(CHAMPCONF)).datatype := datatype;
     end;
     if  (CHAMPTABLE = '') then
     begin
       { rab opérateur }
       THValcomboBox(GetControl(FindEtReplace(CHAMPCONF,'YAD_VAL','YAD_OPER',false))).itemindex := -1;
       { rab lien si pas 3 eme ligne }
       if CHAMPCONF <> 'YAD_VAL3' then
         THValcomboBox(GetControl(FindEtReplace(CHAMPCONF,'YAD_VAL','YAD_LIEN',false))).itemindex := -1;
       THValcomboBox(GetControl(CHAMPCONF)).text := '';
     end;
   end else result := datatype ;
end;

procedure TOM_YALERTES.chargeComboChamps( stPrefixe,ListeDhControle,TypeChamp : string;ListeChamp,ListeLibelle: HTStrings);
var iTable,iChamp,i :integer;
    ListeInterm: HTStringList;
    stInterm,stLib,Prefixe,ListePrefixe,stLib2 : hString;
    TobTable : tob;
  Mcd : IMCDServiceCOM;
  Table     : ITableCOM ;
  FieldList : IEnumerator ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();

  ListeInterm := HTStringList.Create;
  ListePrefixe:=stPrefixe;

  { prise en compte des prefixes des tables liées }
  TobTable:=VH_EntPgi.TobTablesLiees.FindFirst(['YTB_PREFIXEYTA'],[stPrefixe], false);
  While Assigned(TobTable) Do
    begin
    ListePrefixe:=ListePrefixe+';'+TobTable.GetString('YTB_PREFIXEYTB');
    TobTable := VH_EntPgi.TobTablesLiees.FindNext(['YTB_PREFIXEYTA'],[stPrefixe], false);
    end;
  {TobTable:=VH_EntPgi.TobTableAlertes.FindFirst(['YTA_PREFIXE'],[stPrefixe], false);
  if Assigned(TobTable) then
    ListePrefixe:=stPrefixe+';'+TobTable.GetString('YTA_TABLESLIEES');}
  Repeat
    Prefixe:=ReadToKenSt(ListePrefixe);
    if Prefixe <> '' then
      begin

      table := Mcd.getTable(mcd.PrefixetoTable('T'));
      FieldList := Table.Fields;
      FieldList.Reset();
      While FieldList.MoveNext do
        begin
        { en attendant que le problème sur ce champ soit corrigé
        if V_PGI.Dechamps[iTable,iChamp].nom = 'GA_COLLECTION' then continue;}
        if (FieldList.Current as IFieldCOM).libelle = '' then continue;
        if (Pos(ListeDhControle,(FieldList.Current as IFieldCOM).Control)>0) and ((FieldList.Current as IFieldCOM).libelle[1]<>'.') then
            if (TypeChamp<>'D') or ((TypeChamp='D') and ((FieldList.Current as IFieldCOM).tipe='DATE')) then
              ListeInterm.add ((FieldList.Current as IFieldCOM).libelle+';'+(FieldList.Current as IFieldCOM).name);
        end;
      end;
  until Prefixe='';

  ListeInterm.sort;
  for iChamp := 0 to ListeInterm.count-1 do
    begin
    stInterm:=ListeInterm.Strings[iChamp];
    stLib:=ReadToKenSt(stInterm);
    stLib2:=ReadToKenSt(stInterm);
    ListeLibelle.add(stLib{+'('+stLib2+')'});
    ListeChamp.add(stLib2);
    end;
  ListeInterm.free;
end;

procedure TOM_YALERTES.ChargeListeEvent;
var SCB:TScrollBox;
    ListeEvent : hTStringList;
    StChamp : array[1..3] of hstring;
    P:TPanel;
    ii,iElement,indice,Pos1,i : integer;
    StString : hstring;
{$IFNDEF EAGLCLIENT}
    CC:THDBCheckBox;
{$ELSE EAGLCLIENT}
    CC:THCheckBox;
{$ENDIF !EAGLCLIENT}
begin
  // Création de la ScrollBox associée
  SCB:=TScrollBox.create(Ecran) ;
  SCB.Parent:=TGroupBox(getcontrol('EVENEMENTS')) ;
  SCB.ParentFont:=False;
  SCB.Align:=alClient;
  NbEvent:=0;
  ii:=TTToNum(stComboEvent) ;
  if ii>0 then
    begin
    if (V_PGI.DECombos[ii].Valeurs = Nil) then RemplirListe (stComboEvent,'');
    if (V_PGI.DECombos[ii].Valeurs = Nil) then
      begin
      PgiBox (TraduireMemoire('Aucun évènement n''est paramétré. '+#13#10+
         'Vous ne pouvez pas utilisé cette fonctionnalité.'),'Paramétrage des alertes');
      exit;
      end
    end
  else
    begin
    PgiBox (TraduireMemoire('Aucun évènement n''est paramétré. '+#13#10+
         'Vous ne pouvez pas utilisé cette fonctionnalité.'),'Paramétrage des alertes');
    end;

    ListeEvent := hTStringList.Create;
    for i:=0 to Pred(V_PGI.DECombos[ii].Valeurs.count) do ListeEvent.Add (V_PGI.DECombos[ii].Valeurs.Strings[i]);

    NbEvent:=ListeEvent.count;
    iElement := ListeEvent.count - 1;
    While iElement >= 0 do
       begin
       StString := ListeEvent.Strings[iElement];
       for indice := 0 to 2 do
           begin
           Pos1 := Pos (#9,StString);
           if Pos1 > 1 then
              begin
              StChamp[indice+1] := Copy(StString,1,Pos1-1);Delete(StString,1,Pos1);
              end
           else
              begin
              StChamp[indice+1] := '';
              end;
           end;
       P:=TPanel.Create(Ecran);
       if Assigned(P) then
         begin
         P.Name:='P'+stNomChpEvent+copy(StChamp[1],1,1);
         P.Caption:='';
         P.BevelOuter:=bvNone;
         P.Parent:=SCB;
         P.ParentFont:=False;
         P.ParentColor:=true;
         P.Height:=19;
         P.Align:=alTop;
         end
       else
        PgiBox ('Problème création Panel','Paramétrage des alertes');
{$IFNDEF EAGLCLIENT}
       CC:=THDBCheckBox.Create(Ecran);
{$ELSE EAGLCLIENT}
       CC:=THCheckBox.Create(Ecran);
{$ENDIF !EAGLCLIENT}
       if Assigned(CC) then
         begin
         CC.Parent:=P;
         CC.Name:=stNomChpEvent+copy(StChamp[1],1,1);
{$IFNDEF EAGLCLIENT}
         CC.DataField:=CC.Name;
         CC.ValueChecked:='X';
         CC.ValueUnchecked:='-';
         CC.DataSource:=THDBEdit(GetControl('YAL_LIBELLE')).DataSource;
{$ENDIF !EAGLCLIENT}
         CC.Left:=2; CC.Width:=P.Width-6;
         CC.Alignment:= taRightJustify;
         CC.Top:=2;
         CC.Caption:=StChamp[2];
         CC.Anchors:=[akLeft];
         end
       else
        PgiBox ('Problème création Boite à cocher','Paramétrage des alertes');

       Dec (iElement)
       end;
  { modif des taborder qui sont par défaut inversés }
  for i:=0 to NbEvent - 1 Do
    begin
    if Assigned(GetControl('P'+stNomChpEvent+IntToStr(i+1))) then
       TPanel(GetControl('P'+stNomChpEvent+IntToStr(i+1))).TabOrder:=i;
    end;
  ListeEvent.free;
end;

function TOM_YALERTES.supp_suffixe(stName,Suffixe : String) : String;
var i : integer;
begin
  result := '';
  if Suffixe = '' then result:=stName
  else
  for i := length(stName) downto 1 do
    if copy(stName,i, length(stName)-i+1) = Suffixe then
      begin
      result:=copy(stName,1,i-1);
      break;
      end ;
  if result = '' then result:=stName;
end;

Procedure TOM_YALERTES.DuplicationClick( Sender : tObject);
var TobAlerte : tob;
    num1,num2 : integer;
begin
  if (Ecran = Nil) or ( not(Ecran is TFFiche) ) then exit;
  TobAlerte:=TOB.Create ('YALERTES', Nil, -1);
  TobAlerte.GetEcran (TFfiche(Ecran),Nil);
  oldCodeAlerte:=GetField('YAL_ALERTE');
  bDuplication  := true;
  if  (Ecran <> Nil) and ( Ecran is TFFiche ) then
    if not TFFiche (Ecran).Bouge (TNavigateBtn(NbInsert)) then exit;

  TobAlerte.PutEcran(TFfiche(Ecran));

  if TobAlerte.GetString('YAL_PREDEFINI') = 'CEG' then
    FicheReadOnly(Ecran,false);

  SetFocusControl('YAL_ALERTE');
  SetControlEnabled('BDUPLICATION',false);

  if not Isnumeric(copy(V_PGI.NumVersion,1,1)) or not Isnumeric(copy(V_PGI.NumVersion,3,1)) then
    SetField('YAL_VERSIONALERTE','08.000.001')
  else
    begin
    num1:=ValeurI(copy(V_PGI.NumVersion,1,1));
    num2:=ValeurI(copy(V_PGI.NumVersion,3,1));
    SetField('YAL_VERSIONALERTE',Format('%.2d',[num1])+'.'+Format('%.3d',[num2])+'.'+'001');
    end;
  SetField ('YAL_PREFIXE',TobAlerte.GetString('YAL_PREFIXE'));
  SetField ('YAL_DESTINATAIRES',TobAlerte.GetString('YAL_DESTINATAIRES'));
  SetField ('YAL_VALIDATION',TobAlerte.GetString('YAL_VALIDATION'));
  TobAlerte.free;

  ChargeListeConditions(GetField ('YAL_PREFIXE'),'');
  if TCheckBox(GetControl('YAL_EVENEMENT4')).Checked then
    AlimListesChamps(false,'C');

  stAction:='MODIFICATION';
  SetBoutonEnabled(true);
  adapteValidDesti;
  { cas de la duplic d'une alerte CEGID }
  if GetControlText('YAL_PREDEFINI') = 'CEG' then
    setField ('YAL_PREDEFINI','');
  setField ('YAL_ALERTE','');
end;

procedure TOM_YALERTES.adapteValidDesti;
begin
  { Validation simple ou J'ai lu }
  if GetField ('YAL_VALIDATION') <> 'S' then
     SetControlChecked('TYPE_JLOK',true)
  else
     SetControlChecked('TYPE_OK',true);
  { Alerte simple ou croisée}
  { Destinaires Tous, utilisateurs ou groupes }
  AlimTobDestinataires();
  if GetField ('YAL_DESTINATAIRES') = 'U' then
     SetControlChecked('UT_UTIL',true)
  else
     if GetField ('YAL_DESTINATAIRES') = 'G' then
        SetControlChecked('UT_GRP',true)
     else
        begin
        SetControlVisible  ('LIST_UTIL',false);
        SetControlChecked('UT_TOUS',true);
        end;

  if ( ( not AfterInserting ) or (bDuplication) ) and ( GetField ('YAL_DESTINATAIRES') <> '' ) then
     ChargeListUser();
end;

Procedure TOM_YALERTES.SetBoutonEnabled( bb : boolean);
begin
  SetControlEnabled('PAREMAIL',bb);
  SetControlEnabled('BAJOUTER',bb);
  SetControlEnabled('BENLEVER',bb);
end;

Procedure TOM_YALERTES.SetEventEnabled( numEvent : integer);
var i : integer;
begin
  for i:=1 to NbEvent do
    if i <> numEvent then
      begin
      if Assigned(GetControl('YAL_EVENEMENT'+IntToSTr(i))) then
        begin
        TCheckBox(GetControl('YAL_EVENEMENT'+IntToSTr(i))).Checked := false;
        SetField ('YAL_EVENEMENT'+IntToSTr(i),'-')
        end;
      end;
  SetControlVisible  ('SELECTIONCHAMP',(numEvent = 5) or (numEvent= 4));
  SetControlVisible  ('TSCONDITIONS',(numEvent<>5) and (numEvent<>4));
  SetControlVisible  ('YAL_NBJOURS',(numEvent=5));
  SetControlVisible  ('TYAL_NBJOURS',(numEvent=5));
end;

Initialization
  registerclasses ( [ TOM_YALERTES ] ) ;
end.
