{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 06/02/2002
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : INTERIMAIRES (INTERIMAIRES)
Mots clefs ... : TOM;INTERIMAIRES
*****************************************************************
PT1   : 22/05/2002 JL V582 Ajout contrôle NumSS par rapport à date de naissance
PT2   : 27/05/2002 JL V582 Accès aux carte de séjour si nationalité <> FRA
PT3   : 10/10/2002 VG V585 Fonction TestNumeroSS déplacée dans PgOutils
PT4   : 17/12/2002 PH V591 Toutes les dates sont initialisées à idate1900 ou 2099 au lieu de null
PT5   : 26/01/2004 JL V_50 FQ 10697 Ajout nom de jeune fille
PT5   : 10/02/2004 JL V_50 Maj nouveau champs intérimaires gérés pour RH Compétences
}
Unit UtomInterimaires ;

Interface

Uses
     {$IFDEF VER150}
     Variants,
     {$ENDIF}
     StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fiche,HDB,FichList,FE_Main,
{$ELSE}
      eFiche,eFichList,MaineAgl,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOM,UTob,HTB97,PgOutils,PGOutils2,EntPaie;

Type
  TOM_INTERIMAIRES = Class (TOM)
    procedure OnLoadRecord               ; override ;
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    private
    TypeSaisie,Action,NumSS,MoisSS,AnneeSS:String;
    ResultAnnee,ResultMois:Integer;
    procedure BEmploiClick(Sender:TObject);
//    Function  TestNumeroSS( numss: String ; codesexe : String  ) : integer;
    procedure MesErrorSecu( resultat : Integer ) ;
    Function  TestNumeroSSNaissance( numss: String;SSNaiss:String ) : integer; // PT1
    procedure MesErrorSecuNaissance( resultAnnee,ResultMois : Integer ) ;  // PT1
    end ;

Implementation

procedure TOM_INTERIMAIRES.OnLoadRecord ;
Begin
  Inherited ;
If AfterInserting Then Action:='NOUVEAU'
Else Action:='MODIFICATION';
end;

procedure TOM_INTERIMAIRES.OnNewRecord ;
var MaxSal,MaxInt:Integer;
    Q:TQuery;
begin
  Inherited ;
If TypeSaisie = 'CANDIDATS' then SetField('PSI_TYPEINTERIM','CAN');
If TypeSaisie = 'EXTERIEUR' then SetField('PSI_TYPEINTERIM','EXT');
SetControlVisible('BEMPLOI',False);
SetField('PSI_CONFIDENTIEL','-');
// PT4 17/12/2002 PH V591 Toutes les dates sont initialisées à idate1900 ou 2099 au lieu de null
SetField('PSI_DATENAISSANCE',IDate1900);
SetField('PSI_DATEXPIRSEJOUR',IDate1900);
MaxInt:=0;
MaxSal:=0;
if (VH_PAIE.PgTypeNumSal='NUM') and (VH_Paie.PGIncSalarie=TRUE) then
   begin
   if VH_PAIE.PGCodeInterim=True Then
      begin
      MaxSal:=0;
      Q:=OpenSQL('SELECT MAX(PSA_SALARIE) FROM SALARIES',TRUE) ;
      if NOT Q.Eof then
         begin
         If Q.Fields[0].AsString<>'' Then MaxSal := StrToInt (Q.Fields[0].AsString)
         Else MaxSal:=0;
         end;
      Ferme(Q);
      end;
   Q:=OpenSQL('SELECT MAX(PSI_INTERIMAIRE) FROM INTERIMAIRES',True);
   if NOT Q.Eof then
      begin
      If Q.Fields[0].AsString<>'' Then MaxInt := StrToInt (Q.Fields[0].AsString)
      Else MaxInt:=0;
      end;
   Ferme(Q);
   If MaxSal>MaxInt Then MaxInt:=MaxSal;
   SetField ('PSI_INTERIMAIRE',IntToStr (MaxInt + 1));
   end;
end ;

procedure TOM_INTERIMAIRES.OnDeleteRecord;
begin
If ExisteSQL('SELECT PEI_INTERIMAIRE FROM EMPLOIINTERIM WHERE PEI_INTERIMAIRE="'+GetField('PSI_INTERIMAIRE')+'"') Then
   begin
   Case PGIAskCancel('Attention, il existe des périodes d''emploi pour cet intérimaire (ou stagiaire),#13#10 voulez-vous les supprimer aussi ?',TFFiche(Ecran).Caption) of
   mrYes :
         begin
         ExecuteSQL('DELETE FROM EMPLOIINTERIM WHERE PEI_INTERIMAIRE="'+GetField('PSI_INTERIMAIRE')+'"');
         end;
   mrCancel : exit;
   end;
   end;
end;

procedure TOM_INTERIMAIRES.OnUpdateRecord ;
begin
  Inherited ;
If TFFiche(Ecran).FTypeAction=TaCreat Then
   begin
   If (VH_Paie.PGcodeInterim=True) and (VH_Paie.PGInterimaires=True) Then
      begin
      If ExisteSQL('SELECT PSA_SALARIE FROM SALARIES WHERE PSA_SALARIE="'+GetField('PSI_INTERIMAIRE')+'"') Then
         begin
         LastError:=1;
         PgiBox('Création impossible car ce code est déja affecté à un salarié',TFFiche(Ecran).Caption);
         SetFocusControl('PSI_INTERIMAIRE');
         Exit;
         end;
      end;
   end;
If GetField('PSI_TYPEINTERIM')='' Then
   begin
   LastError:=1;
   PgiBox('Vous devez renseigner le type',TFFiche(Ecran).Caption);
   SetFocusControl('PSI_TYPEINTERIM');
   end;
end ;

procedure TOM_INTERIMAIRES.OnChangeField ( F: TField ) ;
var interimaire,TypeInterimaire,CodeSexe:String;
    {$IFNDEF EAGLCLIENT}
    sexe:THDBValComboBox;
    {$ELSE}
    sexe:THValComboBox;
    {$ENDIF}
    resultat:Integer;
begin
  Inherited ;
if (F.FieldName='PSI_LIBELLE') OR (F.FieldName='PSI_PRENOM') then
   begin
   Interimaire:=GetField('PSI_INTERIMAIRE');
   TypeInterimaire:=GetField('PSI_TYPEINTERIM');
   If Interimaire<>'' Then TFFiche(Ecran).Caption:='Saisie '+RechDom('PGTYPEINTERIM',TypeInterimaire,False)+' : '+
                      Interimaire+' '+GetField('PSI_LIBELLE')+' '+GetField('PSI_PRENOM')
   Else TFFiche(Ecran).Caption:='Saisie Intérimaires et stagiaires';
   end;
If F.FieldName='PSI_TYPEINTERIM' Then
   begin
   If (GetField('PSI_TYPEINTERIM')='INT') OR (GetField('PSI_TYPEINTERIM')='STA') Then SetControlVisible('BEMPLOI',True)
   Else SetControlVisible('BEMPLOI',False);
   end;
If (F.FieldName = 'PSI_SEXE') then
  begin
        {$IFNDEF EAGLCLIENT}
        sexe:=THDBValComboBox(GetControl('PSI_SEXE'));
        {$ELSE}
        sexe:=THValComboBox(GetControl('PSI_SEXE'));
        {$ENDIF}
        if (sexe<>nil) then
        begin
                //DEBUT PT5
                If (GetField('PSI_SEXE')='M') and (Getfield('PSI_NOMJF')<>'') then
                begin
                        HShowMessage('6;Incohérence '+Sexe.text+' :'+
                        ';Incohérence Nom de jeune fille et sexe!;E;O;O;O;;;','','');
                        Setfield('PSA_NOMJF','');
                end;
                //FIN PT5
                codesexe :=sexe.value;
                numss :=(GetField('PSI_NUMEROSS'));
                if (codesexe<>'') AND (numss<>'') then
                begin
                        resultat:=TestNumeroSS(numss,codesexe);
                        MesErrorSecu(resultat);
                end;
        end;
  end;
If (F.FieldName = 'PSI_NUMEROSS')  then
  begin
{$IFNDEF EAGLCLIENT}
  sexe:=THDBValComboBox(GetControl('PSI_SEXE'));
{$ELSE}
  sexe:=THValComboBox(GetControl('PSI_SEXE'));
{$ENDIF}
  if sexe<>nil then codesexe :=sexe.value;
  numss :=(GetField('PSI_NUMEROSS'));
  if (codesexe<>'') AND (numss<>'') then
    begin
    If (Length(numss)=15) or (Length(numss)=13)  Then    // PT1  Les test sur l'année, le mois de naissance
       begin                    //sont effectués que si le numéro SS fait 15 caractères, et avant le test de la clé
       resultAnnee:=TestNumeroSSNaissance(numss,'Annee');
       resultMois:=TestNumeroSSNaissance(numss,'Mois');
       MesErrorSecuNaissance(resultAnnee,ResultMois);
       If (ResultAnnee<>0) or (ResultMois<>0) Then  numss :=(GetField('PSI_NUMEROSS'));
       end;
    resultat:=TestNumeroSS(numss,codesexe);
    MesErrorSecu(resultat);
    end;
  exit;
  end;
if (F.FieldName='PSI_INTERIMAIRE') then
  begin
  if (GetField('PSI_INTERIMAIRE'))<>'' then
    begin
    Interimaire:=Trim(GetField('PSI_INTERIMAIRE'));
    if (isnumeric(Interimaire)) and (interimaire <> '') then
      if (StrToInt(interimaire)>2147483647) then
        Begin
        PgiBox('Vous ne pouvez saisir d''intérimaire superieur à 2147483647!','Matricule de l''intérimaire');
        SetField('PSI_INTERIMAIRE','2147483647');
        end;
      if (VH_PAIE.PgTypeNumSal='NUM') and (isnumeric(Interimaire)) and (Interimaire <> '') then
        Begin
        Interimaire := ColleZeroDevant(StrToInt(trim(Interimaire)),10);
        if interimaire <> (GetField('PSI_INTERIMAIRE')) then SetField('PSI_INTERIMAIRE',Interimaire);
        End
      else
        if (VH_PAIE.PgTypeNumSal='ALP') AND (Interimaire<>GetField('PSI_INTERIMAIRE')) then SetField('PSI_INTERIMAIRE',Interimaire);
      end;
  end;
if (F.FieldName=('PSI_DATENAISSANCE')) then
   begin
   numss :=(GetField('PSI_NUMEROSS'));
   If (Length(numss)=15) or(Length(numss)=13) Then
      begin
      resultAnnee:=TestNumeroSSNaissance(numss,'Annee');
      resultMois:=TestNumeroSSNaissance(numss,'Mois');
      MesErrorSecuNaissance(resultAnnee,ResultMois);
              end;
   end;
If (F.FieldName=('PSI_NATIONALITE')) then                //PT2
   begin
   If GetField('PSI_NATIONALITE')<>'FRA' then
      begin
      SetControlEnabled('PSI_CARTESEJOUR',True);
      SetControlEnabled('PSI_DATEXPIRSEJOUR',True);
      SetControlEnabled('PSI_DELIVPAR',True);
      end
   Else
      begin
      SetControlEnabled('PSI_CARTESEJOUR',false);
      SetControlEnabled('PSI_DATEXPIRSEJOUR',false);
      SetControlEnabled('PSI_DELIVPAR',false);
      end;
   end;

   //DEBUT PT5
   if (F.FieldName='PSI_NOMJF') then
   begin
        {$IFNDEF EAGLCLIENT}
        sexe:=THDBValComboBox(GetControl('PSI_SEXE'));
        {$ELSE}
        sexe:=THValComboBox(GetControl('PSI_SEXE'));
        {$ENDIF}
        if sexe <> NIL then
        begin
                If (sexe.value='M') and (Getfield('PSI_NOMJF')<>'') then
                HShowMessage('6;Incohérence '+Sexe.text+' :'+
                ';Incohérence Nom de jeune fille,sexe !;E;O;O;O;;;','','');
        end;
   end;
   // FIN PT5
end;

procedure TOM_INTERIMAIRES.OnArgument ( S: String ) ;
var BEmploi:TToolBarButton97;
{$IFNDEF EAGLCLIENT}
Combo,ComboLieu : THDBValComboBox ;
Edit:THDBEdit;
{$ELSE}
Edit:THEdit;
Combo,ComboLieu : THValComboBox ;
{$ENDIF}
Num : integer;
TCodeStat,TLieu : THLabel;
Numero,Libelle : String;
begin
  Inherited ;
TypeSaisie := ReadTokenPipe(S,';');
If TypeSaisie = 'CANDIDATS' then
begin
        SetControlVisible('BEMPLOI',False);
        SetControlEnabled('PEMPLOI',True);
        SetControlCaption('PEMPLOI','Affectation');
        SetControlProperty('PSI_TYPEINTERIM','Plus','');
        SetControlEnabled('PSI_TYPEINTERIM',False);
        TFFiche(Ecran).Caption := 'Saisie candiats :';
        UpdateCaption(TFFiche(Ecran));
end;
If TypeSaisie = 'EXTERIEUR' then
begin
        SetControlVisible('BEMPLOI',False);
        SetControlEnabled('PEMPLOI',True);
        SetControlCaption('PEMPLOI','Affectation');
        SetControlProperty('PSI_TYPEINTERIM','Plus','');
        SetControlEnabled('PSI_TYPEINTERIM',False);
        TFFiche(Ecran).Caption := 'Intervenant extérieur :';
        UpdateCaption(TFFiche(Ecran));
end;
BEmploi:=TToolBarButton97(GetControl('BEMPLOI'));
If BEmploi<>Nil Then BEmploi.OnClick:=BEmploiClick;
{$IFNDEF EAGLCLIENT}
Edit:=THDBEdit(GetControl('PSI_INTERIMAIRE'));
{$ELSE}
Edit:=THEdit(GetControl('PSI_INTERIMAIRE'));
{$ENDIF}
if (VH_PAIE.PgTypeNumSal='ALP') then
    if Edit<>nil then Edit.EditMask:='';
if (VH_PAIE.PgTypeNumSal='NUM') then
  if Edit<>nil then Edit.EditMask:='9999999999';

//Visiblilité Champ Organisation
For num := 1 to VH_Paie.PGNbreStatOrg do
  begin
  Numero:=InttoStr(num);
  if Num > VH_Paie.PGNbreStatOrg then break;
  libelle:='';
  if Num = 1 then libelle:=VH_Paie.PGLibelleOrgStat1;
  if Num = 2 then libelle :=VH_Paie.PGLibelleOrgStat2;
  if Num = 3 then libelle :=VH_Paie.PGLibelleOrgStat3;
  if Num = 4 then libelle :=VH_Paie.PGLibelleOrgStat4;
  if (Libelle<>'') then
    Begin
{$IFNDEF EAGLCLIENT}
    ComboLieu:=THDBValComboBox(GetControl('PSI_TRAVAILN'+Numero));
{$ELSE}
    ComboLieu:=THValComboBox(GetControl('PSI_TRAVAILN'+Numero));
{$ENDIF}
    if (ComboLieu <> NIL) then ComboLieu.Visible :=TRUE;
    TLieu:=THLabel(GetControl('TPSI_TRAVAILN'+Numero));
    if (TLieu <> NIL)   then
      begin
      TLieu.Visible :=TRUE;
      if Num = 1 then Begin TLieu.Caption :=VH_Paie.PGLibelleOrgStat1; end;
      if Num = 2 then Begin TLieu.Caption :=VH_Paie.PGLibelleOrgStat2; end;
      if Num = 3 then Begin TLieu.Caption :=VH_Paie.PGLibelleOrgStat3; end;
      if Num = 4 then Begin TLieu.Caption :=VH_Paie.PGLibelleOrgStat4; end;
      end;
    End; // begin libelle
  end;
  //Visibilité du code statistique
  if VH_Paie.PGLibCodeStat<>'' then
    begin
     TCodeStat:=THLabel(GetControl('TPSI_CODESTAT'));
{$IFNDEF EAGLCLIENT}
     ComboLieu:=THDBValComboBox(GetControl('PSI_CODESTAT'));
{$ELSE}
     ComboLieu:=THValComboBox(GetControl('PSI_CODESTAT'));
{$ENDIF}
     if TCodeStat<>nil then begin TCodeStat.Caption :=VH_Paie.PGLibCodeStat; TCodeStat.visible:=True; end;
     if ComboLieu<>nil then ComboLieu.Visible :=TRUE;
    end;
For num := 1 to VH_Paie.PgNbCombo do
  begin
  Numero:=InttoStr(num);
  if Num > VH_Paie.PgNbCombo then break;
{$IFNDEF EAGLCLIENT}
  Combo:=THDbValComboBox(GetControl('PSI_LIBREPCMB'+Numero));
{$ELSE}
  Combo:=THValComboBox(GetControl('PSI_LIBREPCMB'+Numero));
{$ENDIF}
  if Combo <> NIL then Combo.Visible :=TRUE;
  TLieu:=THLabel(GetControl('TPSI_LIBREPCMB'+Numero));
  if (TLieu <> NIL)   then
   begin
   TLieu.Visible :=TRUE;
   if Num = 1 then Begin TLieu.Caption :=VH_Paie.PgLibCombo1; end;
   if Num = 2 then Begin TLieu.Caption :=VH_Paie.PgLibCombo2; end;
   if Num = 3 then Begin TLieu.Caption :=VH_Paie.PgLibCombo3; end;
   if Num = 4 then Begin TLieu.Caption :=VH_Paie.PgLibCombo4; end;
   end;
  end;
end ;

procedure TOM_INTERIMAIRES.BEmploiClick(Sender:TObject);
begin
AGLLanceFiche ('PAY', 'EMPLOIINTERIM_MUL','','', GetField('PSI_INTERIMAIRE')+';FICHE');
RefreshDB;
end;

procedure TOM_INTERIMAIRES.MesErrorSecu(resultat : integer);
var cle2,vide,MesError : string;
    NumSecuSaisie,CleSexe:string;
    reponse : Word;
begin
NumSecuSaisie:=GetField('PSI_NUMEROSS');
CleSexe:=NumSecuSaisie[1];
vide:='';
if (resultat <> 0)  then       // resultat intialisé dans la fonction  TestNumeroSS
     begin
      Case resultat of
          1 :  HShowMessage('1;Interimaires:;Vous n''avez pas renseigné le numéro de sécurité sociale.'+
                             '#13#10Pensez-y ulterieurement!;W;O;O;O;;;','','');

         -2 :  MesError:= 'ne comporte pas de clé.#13#10#13#10Veuillez la saisir!';

         -3 :  MesError:= 'passé est incomplet.#13#10NB : 15 positions obligatoires';

         -7 :  MesError:='est incorrect " Vous devez saisir une valeur numérique!!"';

         -5 : begin       //incohérence Psi_Sexe=Femme cle sexe=1
                  reponse:=HShowMessage('1;Clé Sexe Erronée :"'+CleSexe+'";Votre clé sexe du numéro de sécurité sociale vaut "2",'+
                                        '#13#10#13#10Voulez-vous la modifier?;Q;YN;Y;N','','');
                  if reponse = 6 then  // 6 : Oui ; 7 : Non
                    begin  numss[1]:='2';   SetField('PSI_NUMEROSS',numss); end;

               end;

          -6 : begin      // incohérence Psi_Sexe=Homme  cle sexe=2
                  reponse:=HShowMessage('1;Clé Sexe Erronée :"'+CleSexe+'";Votre clé sexe du numéro de sécurité sociale vaut "1",#13#10#13#10'+
                                           'Voulez-vous la modifier?;Q;YN;Y;N;;;','','');
                   if reponse = 6 then
                      begin numss[1]:='1'; SetField('PSI_NUMEROSS',numss);  end;
               end;

          -4 : begin     // clé sexe different de 1 et de 2
                   if  (GetField('PSI_SEXE')='M') then
                     begin
                       reponse:=HShowMessage('1;Clé Sexe Erronée :'+CleSexe+';Votre clé sexe vaut 1,#13#10#13#10'+
                                                    'Voulez-vous la modifier?;Q;YN;Y;N;;;','','');
                       if reponse = 6 then
                         begin numss[1]:='1';  SetField('PSI_NUMEROSS',numss); end;
                     end;
                    if  (GetField('PSI_SEXE')='F') then
                      begin
                        reponse:=HShowMessage('1;Clé Sexe Erronée :'+CleSexe+';Votre clé sexe vaut 2,#13#10Voulez-vous la modifier?;Q;YN;Y;N;;;','','');
                        if reponse = 6 then
                            begin numss[1]:='2';   SetField('PSI_NUMEROSS',numss); end;
                      end;
                end;

          -1 : begin       //clé erronnée
                   if cle<10 then
                       begin
                         cle2:=InttoStr(cle); numss[14]:='0'; Numss[15]:=cle2[1];
                       end;
                   if cle>=10 then
                        begin
                         cle2:=InttoStr(cle); numss[14]:=cle2[1]; numss[15]:=cle2[2];
                        end;
                   reponse:=HShowMessage('1;clé erronée :;Votre clé vaut '+cle2+'!! #13#10Voulez-vous la modifier?;Q;YN;Y;N;;;','','');
                   if reponse=6 then SetField('PSI_NUMEROSS',numss);
                end;
      end;  // end du case
      if (resultat = -2) or  (resultat = -3) or (resultat = -7)then
            HShowMessage('3;N° :'+NumSecuSaisie+' ;Le numéro de sécurité sociale '+MesError+''+
                                ';E;O;O;O;;;','','');


      if (resultat = -7) or (resultat = -3) then begin numss:=''; SetField('PSI_NUMEROSS',numss);
                                   SetfocusControl('PSI_NUMEROSS');
                            end;
   end;
  end;

{PT3
Function TOM_INTERIMAIRES.TestNumeroSS( numss: String ; codesexe : String ) : integer;
var
  i,resultat,k,temp,code:integer;
  numss13,numsscle:string;
  entier : Extended;
  DblCle : Double;
  numcle :integer;
  num13 : Extended;
  // num13 : int64;
begin
 k:= length(numss);
 resultat:=0;
 if ((numss<>'') or (numss='000000000000000'))  then
    begin
        if (resultat=0) and (k<15)  then resultat:=-3;
        if (resultat=-3) and (k=13) then  resultat:=-2;
        if (k=15) and (resultat=0) then          //contrôle du code sexe
          begin
            if ((numss[1]<'1') or (numss[1]>'2'))then resultat:=-4; //incohérence entre code sexe saisie et code sexe conforme
            if numss[1]='1' then     //psI_sexe = homme
                begin
                    if (codesexe='F') then resultat:=-5; //Incohérence entre psi_sexe et code sexe
                end
            else
                begin
                   if numss[1]='2' then         //psi_sexe = femme
                      begin
                          if (codesexe='M') then resultat:=-6; //Incohérence entre PSI_sexe et code sexe
                      end;
                end;
          end;
         if (k=15) and (resultat=0) and (numss[6]='2') then  //contrôle du code département Corse
           begin
            if (numss[7]='A') or (numss[7]='B') then resultat:=2;
           end;
        for i:=1 to k do      //contrôle de numéricité du n° de sécu
          begin
            if resultat<>2 then
            begin
                 if numss[i] in ['0'..'9']  then else resultat:=-7;  //non numérique
            end;
          end;
       if (resultat=0) and (k=15) then      //calcul de la clé
          begin
            //Récupération 13 premier char et cle saisie
            numss13:=Copy(numss,1,13);
            if numss13<>'' then
              Begin
              Num13:=valeur(Numss13);
              numsscle:=numss[14]+numss[15];
              val(numsscle,temp,code);
              numcle:=temp;
              entier:=Int(Num13/97);
              DblCle:= 97-(Num13-(entier*97));
              Cle:=StrToInt(FloatToStr(DblCle));
              if (cle<>numcle)  and (numcle>0) then resultat:=-1
              End;
          end;
    end
  else
   if numss='' then resultat:=1;

result:=resultat;
end;
}

// PT22 procédure d'affichage si erreur ds numéro SS pour : année,mois,département de naissance
procedure TOM_INTERIMAIRES.MesErrorSecuNaissance(resultAnnee,ResultMois : integer);     //PT1
var
   NumSecuSaisie:string;
   Mess:String;
   Reponse:Word;
begin
If (ResultAnnee<>-8) and (ResultMois<>-9) then Exit;
Mess:='';
NumSecuSaisie:=GetField('PSI_NUMEROSS');
If resultAnnee=-8 then // resultat intialisé dans la fonction  TestNumeroSSNaissance
 begin     // Année de naissance non valide
 Mess:=Mess+'- Année de naissance erronnée: '+NumSecuSaisie[2]+NumSecuSaisie[3]+'#13#10l''année vaut '+AnneeSS[3]+AnneeSS[4];
 end;

If resultMois=-9 Then  // resultat intialisé dans la fonction  TestNumeroSSNaissance
 begin     // Mois de naissance non valide
 Mess:=Mess+'#13#10- Mois de naissance erronnée: '+NumSecuSaisie[4]+NumSecuSaisie[5]+'#13#10le mois vaut '+MoisSS[1]+MoisSS[2];
 end;

reponse:=HShowMessage('1;N° de Sécurité Sociale; '+Mess+'#13#10#13#10'+
'Voulez-vous modifier le N° de Sécurité Sociale ?;Q;YN;Y;N;;;','','');
if reponse = 6 then
   begin
   If resultAnnee=-8 then begin numss[2]:=AnneeSS[3]; numss[3]:=AnneeSS[4];end;
   If resultMois=-9 Then begin numss[4]:=MoisSS[1]; numss[5]:=MoisSS[2];end;
   ForceUpdate;
   SetField('PSI_NUMEROSS',numss);
   end;
end;

// PT22 fonction de contrôle ds numéro SS pour : année,mois et département de naissance
Function TOM_INTERIMAIRES.TestNumeroSSNaissance( numss: String;SSNaiss : String ) : integer;  //PT1
var
  aa,mm,jj:word;
  DateNaissanceSS:TDateTime;
begin
Result:=0;
If SSNaiss='Mois' Then
   begin
   resultMois:=0;
   if (numss<>'') or (numss='000000000000000') then
      begin
      If (GetField('PSI_DATENAISSANCE')<>NULL) then
         begin
         DateNaissanceSS:=GetField('PSI_DATENAISSANCE');
         If DateNaissanceSS<>idate1900 Then
            begin
            DecodeDate(DateNaissanceSS,aa,mm,jj);
            If mm<10 Then MoisSS:='0' +IntToStr(mm)
            Else MoisSS:=IntToStr(mm);
            If (NumSS[4]<>MoisSS[1]) OR (NumSS[5]<>MoisSS[2]) Then ResultMois:=-9;
            end;
         end;
      end;
   result:=ResultMois;
   end;
If SSNaiss='Annee' Then
   begin
   ResultAnnee:=0;
   if ((numss<>'') or (numss='000000000000000')) then
      begin
      If (GetField('PSI_DATENAISSANCE')<>NULL)  then
         begin
         DateNaissanceSS:=GetField('PSI_DATENAISSANCE');
         If DateNaissanceSS<>idate1900 Then
            begin
            DecodeDate(DateNaissanceSS,aa,mm,jj);
            AnneeSS:=IntToStr(aa);
            If (NumSS[2]<>AnneeSS[3]) OR (NumSS[3]<>AnneeSS[4]) Then ResultAnnee:=-8;
            end;
         end;
      end;
   result:=resultAnnee;
   end;
end;

Initialization
  registerclasses ( [ TOM_INTERIMAIRES ] ) ;
end.


