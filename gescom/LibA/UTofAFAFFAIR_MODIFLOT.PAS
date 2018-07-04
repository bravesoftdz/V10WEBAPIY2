{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 30/08/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : AFAFFAIR_MODIFLOT ()
Mots clefs ... : TOF;AFAFFAIR_MODIFLOT
*****************************************************************}
Unit UTOFAFAFFAIR_MODIFLOT ;

Interface

uses  StdCtrls,Controls,Classes,
{$IFDEF EAGLCLIENT}
      eMul,  Maineagl,
{$ELSE}
      db,dbTables,mul,FE_Main,
{$ENDIF}
      forms,sysutils,GCMZSUtil,Utob,FactUtil, Ent1,
      ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF, Windows, ParamSoc,
      DicoAF,SaisUtil,EntGC,M3FP,  UtilMultrt,Hstatus,AffaireUtil,
      utofAfBaseCodeAffaire,utilpgi,AglInit,UtilGc,TraducAffaire,AffEcheanceUtil;


Type
  TOF_AFAFFAIR_MODIFLOT = Class (TOF_AFBASECODEAFFAIRE)
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit); override;
    private
    Type_Trait:string ;
    procedure ModifParLotDesAffaires;
    procedure ModifAvecTrait;
  end ;
Type
  TOF_AFMODIFAFF_TRAIT= Class (TOF_AFBASECODEAFFAIRE)
    procedure OnArgument (S : String ) ; override ;
  end ;

Procedure AFLanceFiche_Modiflot_Affaire(TypeTrai:string);


Implementation


procedure TOF_AFAFFAIR_MODIFLOT.OnArgument (S : String ) ;
Var stArg:TStringList;
begin
  Inherited ;
  stArg:=TStringList.Create;
  try
   stArg.Text:=S;
   Type_Trait:=stArg.values['TRAIT'];
   finally
   stArg.free;
   end;
  If Type_Trait='TRA' then
    begin
      Ecran.Caption := TraduitGa('Modification en série des affaires avec traitement');
      updatecaption(Ecran);
    end;
end ;

procedure TOF_AFAFFAIR_MODIFLOT.OnClose ;
begin
  Inherited ;
V_PGI.ExtendedFieldSelection:='' ; //mcd 06/10/2003 10912
end ;


/////////////// ModifParLotDesAffaires //////////////
procedure TOF_AFAFFAIR_MODIFLOT.ModifParLotDesAffaires;
Var F : TFMul ;
    Parametrages : String;
    TheModifLot : TO_ModifParLot;
begin
F:=TFMul(Ecran);
if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
  begin
  V_PGI.ExtendedFieldSelection:='' ;   //mcd 06/10/2003 10912
  MessageAlerte('Aucun élément sélectionné');
  exit;
  end;
If Type_Trait='TRA' then
   begin
   Transactions(ModifAvecTrait,1);  // cas modif de zones affaire qui demande un traitement
   exit;
   end;
//if HShowMessage('0;Confirmation;Confirmez vous la modification des affaires?;Q;YN;N;N;','','')<>mrYes then exit;
TheModifLot := TO_ModifParLot.Create;
TheModifLot.F := F.FListe;
TheModifLot.Q := F.Q;
TheModifLot.NatureTiers := '';
TheModifLot.Nature := 'AFF';
TheModifLot.Titre := Ecran.Caption;
TheModifLot.TableName:='AFFAIRE';
TheModifLot.FCode := 'AFF_AFFAIRE';
TheModifLot.FicheAOuvrir := 'AFFAIRE';
V_PGI.ExtendedFieldSelection:='1';
if ctxscot  in V_PGI.PGIContexte  then
    begin
    TheModifLot.FicheAOuvrir := 'MISSION';
    end;

ModifieEnSerie(TheModifLot, Parametrages) ;
V_PGI.ExtendedFieldSelection:='' ;   //mcd 06/10/2003 10912
end;

procedure TOF_AFAFFAIR_MODIFLOT.ModifAvecTrait;
Var TobAff, TobDet: TOB;
    F : TFMul ;
    select,stmp,Sql,Bool : string;
    ret:T_ErrTraiteEnreg;
    Recep : TStringList;
    wi,NbJour:integer;
    TypeModif : T_TypeModifAff;  // pour type de modif des échéances
    NewDateFin,DateDebut,DateFin,DateDebCal : TdateTime;
    Zmont,X : double;
    Bpourcentage : boolean;
    DEV : RDEVISE;
begin
  F:=TFMul(Ecran);
  TobAff:= Tob.Create('AFFAIRE',NIL,-1);
  Typemodif:=[];
  // on sélectionne les champs à mettre à jour
  Select := AGLLanceFiche('AFF','AFMODIFAFF_TRAIT','','','');
   if (Select='') then
    begin
     TobAff.free;
     exit;
    end;
  Sql := 'SELECT AFF_AFFAIRE, AFF_DATEDEBUT, AFF_DATEFIN, AFF_DATEDEBEXER, AFF_DATEFINEXER ';
  Sql := SQl + ', AFF_DATEMODIF, AFF_ETATAFFAIRE, AFF_DATERESIL, AFF_REPRISEACTIV, AFF_DATEFACTLIQUID';
  Sql := Sql + ', AFF_MONTANTECHE ,AFF_MONTANTECHEDEV,  AFF_POURCENTAGE, AFF_TOTALHTGLODEV, AFF_PROFILGENER';
  Sql := Sql + ', AFF_PERIODICITE, AFF_MULTIECHE, AFF_GENERAUTO, AFF_INTERVALGENER, AFF_METHECHEANCE';
  Sql := Sql + ', AFF_DATEDEBGENER, AFF_DATEFINGENER, AFF_TIERS, AFF_SAISIECONTRE, AFF_DEVISE';
  Sql := Sql +' FROM AFFAIRE left join TIERSCOMPL on YTC_TIERS=AFF_TIERS' ;   // mcd 20/01/2003 ajout jointure pour Ok sélection tierscompl
      // mcd 20/01/2003 mise affafftiers dans le mul
  ret:=TraiteEnregMulTable (F, Sql ,'AFF_AFFAIRE', 'AFFAIRE','AFF_AFFAIRE','AFAFFTIERS',TobAff, True);
  If ret <> tteOK then exit ;
  sTmp := StringReplace(Select, ';', chr(VK_RETURN), [rfReplaceAll]);
  Recep := TStringList.Create;
  Recep.Text := sTmp;
  InitMove(TobAff.Detail.count*2,' ');
  wi := 0;
  while (wi < TobAff.Detail.count) and (V_PGI.IoError=oeOk )  do
    Begin      // si modif dans cette fct, voir UtomAffaire si à modifier aussi
    TobDet := TobAff.detail[wi];
    MoveCur(False);
    Dev.Code := TobDEt.getValue('AFF_DEVISE');
    if Dev.Code = '' then DEV.Code := V_PGI.DevisePivot;
    GetInfosDevise(DEV);
    DEV.Taux:= GetTaux(DEV.Code,DEV.DateTaux,V_PGI.DateEntree) ;
    if (recep.Values['ETAT']<>'') then
      begin   // chgmt code etat. si mis à clo, traitement, sinon RAS.
      if (recep.Values['ETAT'] ='CLO') and (TobDEt.GetValue('AFF_ETATAFFAIRE')<>'CLO') then
        begin   // si modif voir TraitAffTermine dans UtomAffaire
        TobDet.putValue('AFF_ETATAFFAIRE',recep.Values['ETAT']);
          // destruction échéance non facturée
        SupEcheancesAffaire (TobDet.getValue('AFF_AFFAIRE'), False,iDate1900);
        NewDateFin :=AjusteDateGenerSurEch (TobDet.getValue('AFF_AFFAIRE'),False,False,iDate1900,TobDet.GetValue('AFF_DATEDEBGENER'));
        TobDEt.putValue('AFF_DATEFINGENER', NewDatefin);
        if (TobDET.GetValue('AFF_DATEFIN') = idate2099) then TobDEt.PutValue('AFF_DATEFIN',V_PGI.DateEntree);
        end
      else TobDet.putValue('AFF_ETATAFFAIRE',recep.Values['ETAT']);
      end ;
    if  (recep.Values['DATEFIN']<>'31/12/2099') or (recep.Values['DATEDEB']<>'01/01/1900')then
      begin    // on met à jour les date de debut et fin mission on verifiant que l'ona tjrs Fin > début. sinon, on ajuste
      Nbjour := TobDet.getValue('AFF_DATEFIN') - TobDEt.getvalue('AFF_DATEDEBUT');
      if  (recep.Values['DATEFIN']<>'31/12/2099') then TobDet.PutValue ('AFF_DATEFIN', StrToDate(recep.Values['DATEFIN'])) ;
      if  (recep.Values['DATEDEB']<>'01/01/1900') then TobDet.PutValue ('AFF_DATEDEBUT', StrToDate(recep.Values['DATEDEB'])) ;
      if TobDet.getValue('AFF_DATEFIN') < TobDEt.getvalue('AFF_DATEDEBUT') then
        begin  // si les 2 dates sont saisie, on ne passe pas là car tester lors de la saisie des dates
        if  (recep.Values['DATEFIN']<>'31/12/2099') then TobDet.PutValue ('AFF_DATEDEBUT', StrToDate(recep.Values['DATEFIN'])-NbJour);
        if  (recep.Values['DATEDEB']<>'01/01/1900') then TobDet.PutValue ('AFF_DATEFIN', StrToDate(recep.Values['DATEDEB'])+NbJour);
        end;
      end ;
    if  (recep.Values['DATEFINEXER']<>'31/12/2099') or (recep.Values['DATEDEBEXER']<>'01/01/1900')then
      begin    // on met à jour les date de debut et fin exercice on verifiant que l'ona tjrs Fin > début. sinon, on ajuste
      Nbjour := TobDet.getValue('AFF_DATEFINEXER') - TobDEt.getvalue('AFF_DATEDEBEXER');
      if  (recep.Values['DATEFINEXER']<>'31/12/2099') then TobDet.PutValue ('AFF_DATEFINEXER', StrToDate(recep.Values['DATEFINEXER'])) ;
      if  (recep.Values['DATEDEBEXER']<>'01/01/1900') then TobDet.PutValue ('AFF_DATEDEBEXER', StrToDate(recep.Values['DATEDEBEXER'])) ;
      if TobDet.getValue('AFF_DATEFINEXER') < TobDEt.getvalue('AFF_DATEDEBEXER') then
        begin  // si les 2 dates sont saisie, on ne passe pas là car tester lors de la saisie des dates
        if  (recep.Values['DATEFINEXER']<>'31/12/2099') then TobDet.PutValue ('AFF_DATEDEBEXER', StrToDate(recep.Values['DATEFINEXER'])-NbJour);
        if  (recep.Values['DATEDEBEXER']<>'01/01/1900') then TobDet.PutValue ('AFF_DATEFINEXER', StrToDate(recep.Values['DATEDEBEXER'])+NbJour);
        end;
      end ;
    if  (recep.Values['METHE']<>'') and (TobDet.GetValue('AFF_GENERAUTO')='CON') then
      begin
      if (recep.Values['METHE']<> TobDet.Getvalue('AFF_METHECHEANCE'))then
        begin
        TobDet.Putvalue('AFF_METHECHEANCE',recep.Values['METHE']);
        TypeModif:=TypeModif+[TmaDate];
        end;
      end ;
    if  (recep.Values['GENER']<>'') then
      begin  //si modif mode de génération, changement dans les echéances non facturées
      if (recep.Values['GENER']<> TobDet.Getvalue('AFF_GENERAUTO'))then
        begin
        TobDet.Putvalue('AFF_GENERAUTO',recep.Values['GENER']);
        TypeModif:=TypeModif+[TmaDate];
        end;
      end ;
    if  (recep.Values['PROFIL']<>'') then
      begin  //si modif mode de génération, changement dans les echéances non facturées
      if (recep.Values['PROFIL']<> TobDet.Getvalue('AFF_PROFILGENER'))then
        begin
        TobDet.Putvalue('AFF_PROFILGENER',recep.Values['PROFIL']);
        TypeModif:=TypeModif+[TmaDate];
        end;
      end ;
    if  (recep.Values['PERIOD']<>'') then
      begin
      if (recep.Values['PERIOD']<> TobDet.Getvalue('AFF_PERIODICITE'))then
        begin  // si modif péridoicité, recalcul les échéances non facturées
        TobDet.Putvalue('AFF_PERIODICITE',recep.Values['PERIOD']);
        TypeModif:=TypeModif+[Tmadate];
        end;
      end ;
    if  (recep.Values['ACTIV']<>'') then
      begin   // modif type d'activité reprise. reporté sur toutes les échéances non facturées
      If TobDet.Getvalue('AFF_REPRISEACTIV') <> recep.Values['ACTIV'] then
        begin
        TobDet.PutValue('AFF_REPRISEACTIV',recep.Values['ACTIV']);
        TypeModif:=TypeModif+[TmaRepAct];
        end;
      end ;
    if  (recep.Values['INTERV']<>'0') then
      begin // modification del'interval pour la péridocitié. regénération des échénaces non facturées
      If TobDet.Getvalue('AFF_INTERVALGENER') <> StrToInt(recep.Values['INTERV']) then
        begin
        TobDet.PutValue('AFF_INTERVALGENER',StrToInt(recep.Values['INTERV']));
        TypeModif:=TypeModif+[TmaDate];
        end;
      end ;
    if  (recep.Values['POUR']<>'0') and (TobDEt.GetValue('AFF_GENERAUTO')='POU') then
      begin  // modif %. ecrit dans les échénaces non facturées
      If TobDet.Getvalue('AFF_POURCENTAGE') <> StrToFloat(recep.Values['POUR']) then
        begin
        TobDet.PutValue('AFF_POURCENTAGE',StrToFloat(recep.Values['POUR']));
        TypeModif:=TypeModif+[TmaDate];
        end;
      end ;
    if  (recep.Values['MTT']<>'0') and (TobDEt.GetValue('AFF_GENERAUTO')<>'POU') then
      begin // modif mtt, ecrit dans échéances si non facturées
      If TobDet.Getvalue('AFF_MONTANTECHEDEV') <> StrToFloat(recep.Values['MTT']) then
        begin
        TobDet.PutValue('AFF_MONTANTECHEDEV',StrToFloat(recep.Values['MTT']));
        If (DEV.Code <> V_PGI.DevisePivot) Then
          BEGIN
          // if VH^.TenueEuro then BEGIN SufF:='CON' ; SufE:='' ; END else BEGIN SufF:='' ; SufE:='CON' ; END ;
          X:=TobDet.GetValue('AFF_MONTANTECHEDEV') ;
          TobDet.PutValue('AFF_MONTANTECHE',DeviseToEuro (X,DEV.Taux, DEV.Quotite)) ;
          end
        else begin
          TobDEt.PutValue('AFF_MONTANTECHE',TobDEt.GetValue('AFF_MONTANTECHEDEV'));//mcd 14/02/03
          end;
        TypeModif:=TypeModif+[TmaDate];
        end;
      end ;
    if (recep.Values['MUL']='1') then
      begin
      if  (recep.Values['MULTI']='1') then bool:='X'
      else Bool:='-';
      if Bool <> TobDet.GetValue('AFF_MULTIECHE') then
       begin // si mise à X, génère les echeance voulues, sinon, efface celles en trop
       TObDet.putvalue('AFF_MULTIECHE',Bool);
       TypeModif:=TypeModif+[TmaDate];
       end;
      end;
    if  (recep.Values['DATERESIL']<>'31/12/2099') and (TobDet.GetValue('AFF_GENERAUTO')='CON') then
      begin // modif date résiliation. Uniquement si Contrat
      If (TobDet.Getvalue('AFF_DATERESIL') <> StrToDate(recep.Values['DATERESIL']))
        and (StrToDate(recep.Values['DATERESIL']) >= TobDet.Getvalue('AFF_DATEDEBUT'))
        and (StrToDate(recep.Values['DATERESIL']) >= TobDet.Getvalue('AFF_DATEDEBGENER')) then
        begin
        TobDet.PutValue('AFF_DATERESIL',StrToDate(recep.Values['DATERESIL']));
        TypeModif:=TypeModif+[TmaDate];
        end;
      end ;
    if  (recep.Values['LIQUID']<>'31/12/2099') then
      begin  // modif date facture liquidative. mise à jour échéance correspondante dans factaff si pas facturée
      If TobDet.Getvalue('AFF_DATEFACTLIQUID') <> StrToDate(recep.Values['LIQUID']) then
        begin
        TobDet.PutValue('AFF_DATEFACTLIQUID',StrToDate(recep.Values['LIQUID']));
        TypeModif:=TypeModif+[TmaDfacLiq];
        end;
      end ;
    if  (recep.Values['FINFAC']<>'31/12/2099') or (recep.Values['DEBFAC']<>'01/01/1900')then
      begin // modif date début et fin fatcuration. ctrl validiité de 2 date + mise à jour factaff
      If TobDet.Getvalue('AFF_DATEFINGENER') <> StrToDate(recep.Values['FINFAC']) then TypeModif:=TypeModif+[TmaDate];
      If TobDet.Getvalue('AFF_DATEDEBGENER') <> StrToDate(recep.Values['DEBFAC']) then TypeModif:=TypeModif+[TmaDate];
      Nbjour := TobDet.getValue('AFF_DATEFINGENER') - TobDEt.getvalue('AFF_DATEDEBGENER');
      if  (recep.Values['FINFAC']<>'31/12/2099') then TobDet.PutValue ('AFF_DATEFINGENER', StrToDate(recep.Values['FINFAC'])) ;
      if  (recep.Values['DEBFAC']<>'01/01/1900') then TobDet.PutValue ('AFF_DATEDEBGENER', StrToDate(recep.Values['DEBFAC'])) ;
      if TobDet.getValue('AFF_DATEFINGENER') < TobDEt.getvalue('AFF_DATEDEBGENER') then
        begin  // si les 2 dates sont saisie, on ne passe pas là car tester lors de la saisie des dates
        if  (recep.Values['FINFAC']<>'31/12/2099') then TobDet.PutValue ('AFF_DATEDEBGENER', StrToDate(recep.Values['FINFAC'])-NbJour);
        if  (recep.Values['DEBFAC']<>'01/01/1900') then TobDet.PutValue ('AFF_DATEFINGENER', StrToDate(recep.Values['DEBFAC'])+NbJour);
        end;
      end ;
    If (TypeModif<>[]) and (TobDEt.getValue('AFF_ETATAFFAIRE') <>'CLO') then
      begin  // il y a des modifs qui concerne factaff. Il faut appeler la fct qui va bien.
             // si modif voir fct CrerEchenace de UtomAffaire
      DateDebCal := iDate1900;
      if VH_GC.GCIfDefCEGID then
        if GetField('AFF_DATECREATION') < strToDate('15/08/2001') then
           begin
           DateDebCal := StrToDate('15/08/2001');
           end;
      DateDebut  := StrToDate (TobDEt.getValue ('AFF_DATEDEBGENER'));
      if ((TobDEt.getValue('AFF_GENERAUTO') = 'POU') or (TobDEt.getValue('AFF_GENERAUTO') = 'POT')) then
         zmont :=double(TobDEt.getValue('AFF_POURCENTAGE'))
      else
         zmont :=double(TobDEt.getValue ('AFF_MONTANTECHEDEV'));
      DateFin    := StrToDate (TobDEt.getValue ('AFF_DATEFINGENER'));
      If (DateFin >= GetParamSoc('SO_AFDATEFINGENER')) then
         begin
         If GetParamSoc('SO_AFDATEFINGENER')<> iDate2099 then
            DateFin := GetParamSoc('SO_AFDATEFINGENER')
         else
            DateFin := PlusDate (DateDebut,1,'A');
         end;
      if (TobDEt.getValue ('AFF_GENERAUTO')='CON') and (TobDEt.getValue ('AFF_METHECHEANCE')='CIV') then
          DateDebut:=DateDebutPeriode(TobDEt.getValue('AFF_PERIODICITE'), DateDebut,Integer(TobDEt.getValue ('AFF_INTERVALGENER')));
      bPourcentage := ((TobDEt.getValue('AFF_GENERAUTO') = 'POU') or (TobDEt.getValue('AFF_GENERAUTO') = 'POT'));
      Try
        UtilCalculEcheances(TobDet.GetValue('AFF_TOTALHTGLODEV'), TobDet.GetValue('AFF_PROFILGENER'),TobDet.GetValue('AFF_AFFAIRE'),
           TobDEt.getValue('AFF_PERIODICITE'), TobDet.GetValue('AFF_REPRISEACTIV'), TobDet.GetValue('AFF_TIERS'),
           TobDet.GetValue('AFF_GENERAUTO'),
           TypeModif, Integer(TobDet.GetValue('AFF_INTERVALGENER')), zmont, DateDebut, DateFin, TobDet.GetValue('AFF_DATEFACTLIQUID'),
           DateDebCal,idate2099, TobDet.GetValue('AFF_DATERESIL'),
           DEV, (TobDet.GetValue('AFF_SAISIECONTRE')= 'X'), True, bPourcentage, TobDet.GetValue('AFF_MULTIECHE') );
        except
        V_PGI.IoError:=oeSaisie;
        end;
       end;
    TObDet.putvalue('AFF_DATEMODIF', V_PGI.DateEntree);
    inc(wi);
    End;  // fin while  ,boucle sur la sélection
  Recep.Free;
  If V_PGI.IoError=OeOk then
    begin
    Try
     TobAff.UpdateDb(False,False);
     except
     V_PGI.IoError:=oeSaisie;
     end;
    end;
  TobAff.free;
  FiniMove;
end ;


// Attention, pas de AFFAIRE0 déclaré : le champ n'existe pas sur la form
procedure TOF_AFAFFAIR_MODIFLOT.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_:THEdit);
begin
Aff:=THEdit(GetControl('AFF_AFFAIRE'));
//Aff0:=THEdit(GetControl('AFF_AFFAIRE0'));
Aff1:=THEdit(GetControl('AFF_AFFAIRE1'));
Aff2:=THEdit(GetControl('AFF_AFFAIRE2'));
Aff3:=THEdit(GetControl('AFF_AFFAIRE3'));
Aff4:=THEdit(GetControl('AFF_AVENANT'));
Tiers:=THEdit(GetControl('AFF_TIERS'));
end;

/////////////// Procedure appeleé par le bouton Validation //////////////
procedure AGLModifParLotDesAffaires(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then TOTOF:=TFMul(F).LaTOF else exit;
if (TOTOF is TOF_AFAFFAIR_MODIFLOT) then TOF_AFAFFAIR_MODIFLOT(TOTOF).ModifParLotDesAffaires else exit;
end;
Procedure AFLanceFiche_Modiflot_Affaire(TypeTRai:string);
begin
AGLLanceFiche('AFF','AFAFFAIR_MODIFLOT','','','TRAIT='+TypeTrai);
end;

procedure TOF_AFMODIFAFF_TRAIT.OnArgument (S : String ) ;
begin
  Inherited ;
  If CtxScot in V_PGI.PGIContexte then
    begin
    SetControlVisible ('DATERESIL',False);
    SetControlVisible ('TDATERESIL',False);
    SetControlVisible ('METHE',False);
    SetControlVisible ('TMETHE',False);
    SetControlProperty('GENER','Plus','GA" AND CO_CODE<>"CON');
    end
  else begin
    SetControlVisible ('DATEDEBEXER',False);
    SetControlVisible ('DATEFINEXER',False);
    SetControlVisible ('TDATEDEBEXER',False);
    SetControlVisible ('TDATEFINEXER',False);
    // C.B 15/10/03 pour enlever les valeurs btp
    // on enleve aussi Contrat ???
    SetControlProperty('GENER','Plus','GA');                    
    end ;

  if (GetParamSoc('SO_AFGERELIQUIDE')=false) then
    begin
    SetControlVisible('LIQUID',False);
    SetControlVisible('TLIQUID',False);
    end;
  if (GetParamSoc ('SO_AFMULTIECHE') = false) then
    begin
      SetControlVisible ('MUL', False);
    end

end ;


Initialization
  registerclasses ( [ TOF_AFAFFAIR_MODIFLOT ] ) ;
  registerclasses ( [ TOF_AFMODIFAFF_TRAIT ] ) ;
  RegisterAglProc('ModifParLotDesAffaires',TRUE,0,AGLModifParLotDesAffaires);

end.
