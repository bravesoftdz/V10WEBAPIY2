unit UTofRTProspect_ModifLot;

interface
uses UTOF,GCMZSUtil,
{$IFDEF EAGLCLIENT}
     eMul,MainEAGL,
{$ELSE}
     Mul,Fe_Main,
{$ENDIF}
{$ifdef AFFAIRE}
      UtofAfTraducChampLibre,
{$ENDIF}
{$IFDEF GIGI}
     EntGc,
{$ENDIF}
     Forms,Classes,M3FP,HMsgBox,Controls,HEnt1,HCtrls,UtilSelection,UtilGc,ParamSoc,EntRT,
     utob;
Type
{$ifdef AFFAIRE}
                //mcd 11/05/2006 12940  pour faire affectation depuis ressource si paramétré
     TOF_RTProspect_ModifLot = Class (TOF_AFTRADUCCHAMPLIBRE)
 {$else}
     TOF_RTProspect_ModifLot = Class(TOF)
{$endif}
       private
       procedure ModificationParLotDesProspects;
       public
       Argument,Origine : string;
       procedure OnLoad ; override ;
       procedure OnClose ; override ;
       procedure OnArgument (stArgument : string); override;
     end;
Const
   { pour pouvoir gérer les particularité propre à chaque fiche
    fiche Client/prospect à part }
   CodeFichierProspect : String = '0';
   CodeFichierFournisseur : String = '3';
   CodeFichierContact : String = '6';

procedure AGLModifParLotDesProspects(parms : array of variant; nb : integer);
Procedure RTLanceFiche_RTProspect_ModifLot(Nat,Cod : String ; Range,Lequel,Argument : string) ;

implementation
uses 
   CbpMCD
   ,CbpEnumerator;

Procedure RTLanceFiche_RTProspect_ModifLot(Nat,Cod : String ; Range,Lequel,Argument : string) ;
var arg,stArgument,code : string;
    TobChampsProFille : tob;
begin
  arg:=Argument;
  stArgument:=ReadTokenSt(arg);
  if (stArgument = 'PRO') or (stArgument = 'ICF') or (stArgument = 'CCO') then
    begin
    if (stArgument = 'PRO') then code:= CodeFichierProspect
    else if (stArgument = 'ICF') then code:= CodeFichierFournisseur
    else if (stArgument = 'CCO') then code:= CodeFichierContact;

    VH_RT.TobChampsPro.Load;

    TobChampsProFille:=VH_RT.TobChampsPro.FindFirst(['CO_CODE'], [code], TRUE);
    if (TobChampsProFille = Nil ) or (TobChampsProFille.detail.count = 0 ) then
        begin
        PGIInfo(TraduireMemoire('Vous ne pouvez pas effectuer ce traitement.#13#10Le paramètrage de cette saisie n''a pas été effectué'),'');
        exit;
        end;
    end;

AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_RTProspect_ModifLot.OnArgument (stArgument : string);
var F : TForm;
    iTable, iChamp : integer;
    ListeChamp,ListeLibelle,ListeChamp1,ListeLibelle1: TStrings;
    Ctrl : String;
    Mcd : IMCDServiceCOM;
    Table     : ITableCOM ;
    FieldList : IEnumerator ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();

{$ifdef AFFAIRE}
fMulDeTraitement  := true;
inherited;	//mcd 11/05/2006 pour passer dans Aftraducaffaire
{$endif}
	 fTableName := 'TIERS';
Argument:=ReadTokenSt(stArgument);
Origine:=ReadTokenSt(stArgument);

F := TForm (Ecran);
//   SetControlProperty('RTQUALITE','DBListe','RTMULQUALITECONTA') ;
(* mcd 26/06/07 10661 GRC ??? pourquoi # sur affaire ??
if not (ctxAffaire in V_PGI.PGIContexte) then
begin *)
if Origine = 'GC' then
if Argument = 'CLI' then Ecran.Caption:=TraduireMemoire('Modification en série des clients')
  else if Argument = 'ICF' then Ecran.Caption:=TraduireMemoire('Modification en série des infos complémentaires fournisseurs')
       else Ecran.Caption:=TraduireMemoire('Modification en série des fournisseurs')
else
if Argument = 'CLI' then Ecran.Caption:=TraduireMemoire('Modification en série des fiches clients/prospects')
  else if Argument = 'PRO' then Ecran.Caption:=TraduireMemoire('Modification en série des infos complémentaires clients/prospects')
       else if Argument = 'CON' then
            begin
              Ecran.Caption:=TraduireMemoire('Modification en série des fiches contacts');
              SetControlProperty('BZOOM','Hint',TraduireMemoire('Zoom sur le contact'));
            end
            else if Argument = 'CCO' then
                begin
                  Ecran.Caption:=TraduireMemoire('Modification en série des infos complémentaires contacts');
                  SetControlProperty('BZOOM','Hint',TraduireMemoire('Zoom sur le contact'));
                end
                else Ecran.Caption:=TraduireMemoire('Modification en série des fiches suspects');
(*  mcd 26/06/07 10661 GRC
  end else
  Ecran.Caption := TraduireMemoire('Modification en série');*)
UpdateCaption(Ecran);

if Argument<>'SUS' then
begin
if (Argument='CON') or (Argument='CCO')
//    then TFMul(F).Q.Liste:='RTMULQUALITECONTA'
    then TFMul(F).SetDBListe('RTMULQUALITECONTA')
    else
      if (Argument='ICF') or (Argument='FOU') then
//         TFMul(F).Q.Liste:='RFMULQUALITE'
         TFMul(F).SetDBListe('RFMULQUALITE')
      else
//         TFMul(F).Q.Liste:='RTMULQUALITE' ;
         TFMul(F).SetDBListe('RTMULQUALITE') ;
end;

if (Origine = 'GRC') or (Origine = 'GC') then
    begin
     if Argument<>'SUS' then
        if (Argument<>'FOU') and (Argument<>'ICF') then MulCreerPagesCL(F,'NOMFIC=GCTIERS')
        else
          if GetParamSocSecur('SO_RTGESTINFOS003',False) = True then
              MulCreerPagesCL(F,'NOMFIC=GCFOURNISSEURS');
    end;

if (Origine <> 'GRC') then
    begin
    // ajout MCD 31/01/01 pour traduction stat si existe
    GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'YTC_TABLELIBRETIERS', 10, '');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_VALLIBRE', 3, '');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_DATELIBRE', 3, '');
    GCMAJChampLibre (TForm (Ecran), False, 'BOOL', 'YTC_BOOLLIBRE', 3, '');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_TEXTELIBRE', 3, '');
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_RESSOURCE', 3, '');
    If not(ctxScot in V_PGI.PGIContexte)  then
        begin
        SetControlVisible ('TT_MOISCLOTURE' , False);
        SetControlVisible ('T_MOISCLOTURE' , False);
        SetControlVisible ('TT_MOISCLOTURE_' , False);
        SetControlVisible ('T_MOISCLOTURE_' , False);
        end;
    //SetControlVisible ('T_NATUREAUXI_', False);
    //SetControlVisible ('TT_NATUREAUXI', False);
    if Argument = 'CLI' then
    begin
       SetControlText ('T_NATUREAUXI', 'CLI');
       SetControlVisible ('T_REPRESENTANT', GereCommercial);
       SetControlVisible ('TT_REPRESENTANT', GereCommercial);
    end else
       begin
       SetControlVisible ('TABLESLIBRES', False);
       SetControlVisible ('ZONESLIBRES', False);
       SetControlText ('T_NATUREAUXI', 'FOU');
       SetControlProperty('T_TIERS','Datatype','GCTIERSFOURN');
       SetControlVisible ('T_ENSEIGNE', False);SetControlVisible ('TT_ENSEIGNE', False);
       SetControlVisible ('T_SOCIETEGROUPE', False);SetControlVisible ('TT_SOCIETEGROUPE', False);
       SetControlVisible ('T_PRESCRIPTEUR', False);SetControlVisible ('TT_PRESCRIPTEUR', False);
       SetControlVisible ('T_REPRESENTANT', False);SetControlVisible ('TT_REPRESENTANT', False);
       SetControlVisible ('T_ZONECOM', False);SetControlVisible ('TT_ZONECOM', False);
       end;
    end;

if ctxAffaire  in V_PGI.PGIContexte  then begin
   if Argument='CLI' then V_PGI.ExtendedFieldSelection:='3';
   if Argument='FOU' then V_PGI.ExtendedFieldSelection:='4';
   end
else begin
   if Argument='CLI' then V_PGI.ExtendedFieldSelection:='2';
   if Argument='FOU' then V_PGI.ExtendedFieldSelection:='1';
   end;

if (Argument='PRO') or (Argument='CON') or (Argument='CCO') or (Argument='ICF') then V_PGI.ExtendedFieldSelection:='Z';

{if Argument='CON' then
   V_PGI.ExtendedFieldSelection:='L';  }

If V_PGI.ExtendedFieldSelection<>'' then Ctrl:=V_PGI.ExtendedFieldSelection else Ctrl := 'Z';

{ je remet à blanc car fait planter le paramétrage liste }
V_PGI.ExtendedFieldSelection:='';

if (F.Name = 'RTQUALITE') or (F.Name = 'RTQUALITE_SUSP') then
   begin
   ListeChamp := TStringList.Create ;
   ListeLibelle := TStringList.Create;
   ListeChamp1 := TStringList.Create ;
   ListeLibelle1 := TStringList.Create;

   ListeChamp.add('<<Aucun>>');
   ListeLibelle.add (TraduireMemoire('<<Aucun>>'));
   ListeChamp1.add('<<Aucun>>');
   ListeLibelle1.add (TraduireMemoire('<<Aucun>>'));
   if (Argument = 'CLI') or (Argument = 'FOU') then
       begin
       // Table Tiers

        table := Mcd.getTable('TIERS');
        FieldList := Table.Fields;
        FieldList.Reset();
        While FieldList.MoveNext do
           begin
           //if pos( ctrl,V_PGI.DEChamps[iTable,iChamp].Control)=0 then continue ;
           // on ne prend plus tous les champs de la table tiers vue que depuis la 5.0, la vue RTTIERS
           // contient une liste de champs limités
//           if Pos((FieldList.Current as IFieldCOM).name,V_PGI.DEVues[VueToNum('RTTIERS')].SQL ) = 0 then continue;

//           if V_PGI.Dechamps[iTable,iChamp].libelle[1] <> '.' then
           if Copy((FieldList.Current as IFieldCOM).libelle,1,1) <> '.' then
              begin
              ListeChamp.add((FieldList.Current as IFieldCOM).name);
              ListeLibelle.add ((FieldList.Current as IFieldCOM).libelle);
              ListeChamp1.add((FieldList.Current as IFieldCOM).name);
              ListeLibelle1.add ((FieldList.Current as IFieldCOM).libelle);
              end;
           end;
        table := Mcd.getTable(mcd.PrefixetoTable('YTC'));
        FieldList := Table.Fields;
        FieldList.Reset();
        While FieldList.MoveNext do
           begin
           if pos( ctrl,(FieldList.Current as IFieldCOM).control)=0 then continue ;

           if Copy((FieldList.Current as IFieldCOM).libelle,1,1) <> '.' then
              begin
              ListeChamp.add((FieldList.Current as IFieldCOM).name);
              ListeLibelle.add ((FieldList.Current as IFieldCOM).libelle);
              ListeChamp1.add((FieldList.Current as IFieldCOM).name);
              ListeLibelle1.add ((FieldList.Current as IFieldCOM).libelle);
              end;
           end;
       end
       else
       if (Argument = 'PRO') then
         begin
        table := Mcd.getTable(mcd.PrefixetoTable('RPR'));
        FieldList := Table.Fields;
        FieldList.Reset();
        While FieldList.MoveNext do
             begin
             if pos( ctrl,(FieldList.Current as IFieldCOM).control)=0 then continue ;

             if Copy((FieldList.Current as IFieldCOM).libelle,1,1) <> '.' then
                begin
                ListeChamp.add((FieldList.Current as IFieldCOM).name);
                ListeLibelle.add ((FieldList.Current as IFieldCOM).libelle);
                ListeChamp1.add((FieldList.Current as IFieldCOM).name);
                ListeLibelle1.add ((FieldList.Current as IFieldCOM).libelle);
                end;
             end;
         end
         else
         if (Argument = 'CON') then
         begin
					table := Mcd.getTable(mcd.PrefixetoTable('C'));
          FieldList := Table.Fields;
          FieldList.Reset();
          While FieldList.MoveNext do
             begin
             //if pos( ctrl,V_PGI.DEChamps[iTable,iChamp].Control)=0 then continue ;

             if Copy((FieldList.Current as IFieldCOM).libelle,1,1) <> '.' then
                begin
                ListeChamp.add((FieldList.Current as IFieldCOM).name);
                ListeLibelle.add ((FieldList.Current as IFieldCOM).libelle);
                ListeChamp1.add((FieldList.Current as IFieldCOM).name);
                ListeLibelle1.add ((FieldList.Current as IFieldCOM).libelle);
                end;
             end;
         end
         else
         if (Argument = 'CCO') then
         begin
					table := Mcd.getTable(mcd.PrefixetoTable('RD6'));
          FieldList := Table.Fields;
          FieldList.Reset();
          While FieldList.MoveNext do
             begin
             //if pos( ctrl,V_PGI.DEChamps[iTable,iChamp].Control)=0 then continue ;

             if Copy((FieldList.Current as IFieldCOM).libelle,1,1) <> '.' then
                begin
                ListeChamp.add((FieldList.Current as IFieldCOM).name);
                ListeLibelle.add ((FieldList.Current as IFieldCOM).libelle);
                ListeChamp1.add((FieldList.Current as IFieldCOM).name);
                ListeLibelle1.add ((FieldList.Current as IFieldCOM).libelle);
                end;
             end;
         end
         else
         if (Argument = 'ICF') then
         begin
					table := Mcd.getTable(mcd.PrefixetoTable('RD3'));
          FieldList := Table.Fields;
          FieldList.Reset();
          While FieldList.MoveNext do
             begin
             //if pos( ctrl,V_PGI.DEChamps[iTable,iChamp].Control)=0 then continue ;

             if Copy((FieldList.Current as IFieldCOM).libelle,1,1) <> '.' then
                begin
                ListeChamp.add((FieldList.Current as IFieldCOM).name);
                ListeLibelle.add ((FieldList.Current as IFieldCOM).libelle);
                ListeChamp1.add((FieldList.Current as IFieldCOM).name);
                ListeLibelle1.add ((FieldList.Current as IFieldCOM).libelle);
                end;
             end;
         end
         else
         begin
					table := Mcd.getTable(mcd.PrefixetoTable('RSU'));
          FieldList := Table.Fields;
          FieldList.Reset();
          While FieldList.MoveNext do
             begin
             if pos( ctrl,(FieldList.Current as IFieldCOM).Control)=0 then continue ;

             if Copy((FieldList.Current as IFieldCOM).libelle,1,1) <> '.' then
                begin
                ListeChamp.add((FieldList.Current as IFieldCOM).name);
                ListeLibelle.add ((FieldList.Current as IFieldCOM).libelle);
                ListeChamp1.add((FieldList.Current as IFieldCOM).name);
                ListeLibelle1.add ((FieldList.Current as IFieldCOM).libelle);
                end;
             end;
					table := Mcd.getTable(mcd.PrefixetoTable('RSC'));
          FieldList := Table.Fields;
          FieldList.Reset();
          While FieldList.MoveNext do
             begin
             if pos( ctrl,(FieldList.Current as IFieldCOM).Control)=0 then continue ;

             if Copy((FieldList.Current as IFieldCOM).libelle,1,1) <> '.' then
                begin
                ListeChamp.add((FieldList.Current as IFieldCOM).name);
                ListeLibelle.add ((FieldList.Current as IFieldCOM).libelle);
                ListeChamp1.add((FieldList.Current as IFieldCOM).name);
                ListeLibelle1.add ((FieldList.Current as IFieldCOM).libelle);
                end;
             end;

         end;

   THValcomboBox(GetControl('LISTECHAMPS')).Items.Assign(ListeLibelle);
   THValcomboBox(GetControl('LISTECHAMPS')).Values.Assign(ListeChamp);
   THValcomboBox(GetControl('LISTECHAMPS1')).Items.Assign(ListeLibelle1);
   THValcomboBox(GetControl('LISTECHAMPS1')).Values.Assign(ListeChamp1);
   ListeChamp.free;ListeLibelle.Free;
   ListeChamp1.free;ListeLibelle1.Free;
   end;
   if ( (Argument = 'CON') or (Argument = 'CCO') ) and (GetParamSocSecur('SO_RTGESTINFOS006',False) = True ) then
    MulCreerPagesCL(F,'NOMFIC=YYCONTACT');
  if (GetControl('YTC_RESSOURCE1') <> nil)  then
    begin
    if not (ctxaffaire in V_PGI.PGICONTEXTE) then SetControlVisible ('PRESSOURCE',false)
    else begin
      GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_RESSOURCE', 3, '_');
      if not (ctxscot in V_PGI.PGICOntexte) then
         begin
         SetControlVisible ('T_MOISCLOTURE',false);
         SetControlVisible ('T_MOISCLOTURE_',false);
         SetControlVisible ('TT_MOISCLOTURE',false);
         SetControlVisible ('TT_MOISCLOTURE_',false);
         end;
      end;
    end;
{$Ifdef GIGI}
 if (GetControl('T_REPRESENTANT') <> nil) then  SetControlVisible('T_REPRESENTANT',false);
 if (GetControl('TT_REPRESENTANT') <> nil) then  SetControlVisible('TT_REPRESENTANT',false);
 if (GetControl('T_ZONECOM') <> nil) then  SetControlVisible('T_ZONECOM',false);
 if (GetControl('TT_ZONECOM') <> nil) then  SetControlVisible('TT_ZONECOM',false);
 if (F.name <> 'RTQUALITE_SUSP') then
   begin
   if (argument<>'FOU')  then  //mcd 27/08/07 14444 utilisé pour FOU, il en faut pas effacer dans ce cas
    begin
    SetControlText('T_NatureAuxi','');    //on efface les valeurs CLI et PO, car NCP en plus
    SetControlProperty ('T_NATUREAUXI', 'Complete', true);
    SetControlProperty ('T_NATUREAUXI_', 'Complete', true);
    SetControlProperty ('T_NATUREAUXI', 'Datatype', 'TTNATTIERS');
    SetControlProperty ('T_NATUREAUXI_', 'Datatype', 'TTNATTIERS');
    SetControlProperty ('T_NATUREAUXI_', 'Plus', VH_GC.AfNatTiersGRCGI);
    SetControlProperty ('T_NATUREAUXI', 'Plus', VH_GC.AfNatTiersGRCGI);
    end;
   end
 else begin
  if (GetControl('RSU_REPRESENTANT') <> nil) then  SetControlVisible('RSU_REPRESENTANT',false);
  if (GetControl('TRSU_REPRESENTANT') <> nil) then  SetControlVisible('TRSU_REPRESENTANT',false);
  if (GetControl('RSU_ZONECOM') <> nil) then  SetControlVisible('RSU_ZONECOM',false);
  if (GetControl('TRSU_ZONECOM') <> nil) then  SetControlVisible('TRSU_ZONECOM',false);
  end;
{$endif}
end;

procedure TOF_RTProspect_ModifLot.OnLoad;
begin
end;

procedure TOF_RTProspect_ModifLot.OnClose;
begin
V_PGI.ExtendedFieldSelection:='' ;
end;

/////////////// ModificationParLotDesProspects //////////////
procedure TOF_RTProspect_ModifLot.ModificationParLotDesProspects;
Var F : TFMul ;
    Parametrages : String;
    TheModifLot : TO_ModifParLot;
begin
{ mng : je remets ce bloc ici car "Paramétrer la liste" remet à blanc V_PGI.ExtendedFieldSelection }
if ctxAffaire  in V_PGI.PGIContexte  then begin
   if Argument='CLI' then V_PGI.ExtendedFieldSelection:='3';
   if Argument='FOU' then V_PGI.ExtendedFieldSelection:='4';
   end
else begin
   if Argument='CLI' then V_PGI.ExtendedFieldSelection:='2';
   if Argument='FOU' then V_PGI.ExtendedFieldSelection:='1';
   end;

if (Argument='PRO') or (Argument='CON') or (Argument='CCO') or (Argument='ICF') then V_PGI.ExtendedFieldSelection:='Z';

{if Argument='CON' then
   V_PGI.ExtendedFieldSelection:='L';  }

F:=TFMul(Ecran);
if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
  begin
  V_PGI.ExtendedFieldSelection:='' ;
  MessageAlerte('Aucun élément sélectionné');
  exit;
  end;
{$IFDEF EAGLCLIENT}
if F.bSelectAll.Down then
   if not F.Fetchlestous then
     begin
     F.bSelectAllClick(Nil);
     F.bSelectAll.Down := False;
     V_PGI.ExtendedFieldSelection:='' ;
     exit;
     end else
     begin
     F.bSelectAll.Down := True;
     F.Fliste.AllSelected := true;
     end;
{$ENDIF}

TheModifLot := TO_ModifParLot.Create;
TheModifLot.F := F.FListe; TheModifLot.Q := F.Q;
TheModifLot.NatureTiers := Argument;
TheModifLot.Titre := Ecran.Caption;
// mng 18-11-03 pour restriction fiche
TheModifLot.stAbrege :='';
if Argument = 'PRO' then
    begin
    TheModifLot.TableName:='PROSPECTS';
    TheModifLot.FCode := 'T_AUXILIAIRE';
    TheModifLot.Nature := 'RT';
    TheModifLot.FicheAOuvrir := 'RTPARAMCL';
    // mng 18-11-03 pour restriction fiche
    TheModifLot.stAbrege:='T_NATUREAUXI';
    end
    else
    if Argument = 'CON' then
        begin
        TheModifLot.TableName:='CONTACT';
        TheModifLot.FCode := 'C_TYPECONTACT;C_AUXILIAIRE;C_NUMEROCONTACT';
        TheModifLot.Nature := 'YY';
        TheModifLot.FicheAOuvrir := 'YYCONTACT';
        end
        else
        if Argument = 'CCO' then
            begin
            TheModifLot.TableName:='RTINFOS006';
            TheModifLot.FCode := 'RD6_CLEDATA';
            TheModifLot.Nature := 'RT';
            TheModifLot.FicheAOuvrir := 'RTPARAMCL';
            end
            else
            if Argument = 'ICF' then
                begin
                TheModifLot.TableName:='RTINFOS003';
                TheModifLot.FCode := 'RD3_CLEDATA';
                TheModifLot.Nature := 'RT';
                TheModifLot.FicheAOuvrir := 'RTPARAMCL';
                end
                else
                if Argument = 'SUS' then
                    begin
                    TheModifLot.TableName:='SUSPECTS;SUSPECTSCOMPL';
                    TheModifLot.FCode := 'RSU_SUSPECT';
                    TheModifLot.Nature := 'RT';
                    TheModifLot.FicheAOuvrir := 'RTSUSPECTS';
                    end
                    else
                    begin
                    TheModifLot.TableName:='TIERS;TIERSCOMPL';
                    TheModifLot.FCode := 'T_AUXILIAIRE';
                    TheModifLot.Nature := 'GC';
                    TheModifLot.FicheAOuvrir := 'GCTIERS';
                    TheModifLot.stAbrege:='T_NATUREAUXI';
                    // mng 18-11-03 pour restriction fiche
                    if Argument='FOU' then TheModifLot.FicheAOuvrir := 'GCFOURNISSEUR';
                    end;

ModifieEnSerie(TheModifLot, Parametrages) ;

if F.bSelectAll.Down then
    begin
    F.bSelectAllClick(Nil);
    F.bSelectAll.Down := False;
    end;
V_PGI.ExtendedFieldSelection:='' ;
end;

/////////////// Procedure appellé par le bouton Validation //////////////
procedure AGLModifParLotDesProspects(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then TOTOF:=TFMul(F).LaTOF else exit;
if (TOTOF is TOF_RTProspect_ModifLot) then TOF_RTProspect_ModifLot(TOTOF).ModificationParLotDesProspects else exit;
end;

Initialization
registerclasses([TOF_RTProspect_ModifLot]) ;
RegisterAglProc('ModifParLotDesProspects',TRUE,0,AGLModifParLotDesProspects);
end.
