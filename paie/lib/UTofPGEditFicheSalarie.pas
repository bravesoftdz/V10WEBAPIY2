{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 27/06/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGEDITFICHESAL ()
Mots clefs ... : TOF;PGEDITFICHESAL
*****************************************************************
PT1   : 21/11/2003 JL V_50 FQ 10543 Ajout édition contrôle des fiches salariés
PT2   : 28/06/2004 JL V_50 FQ 10543 Ajout tri par thème et correction sur champs
                           libres
PT3   : 14/09/2004 JL V_50 FQ 11597 test sur date de sortie était inversé
PT4   : 06/12/2004 JL V_60 FQ 11742 Testb champs null pour V_61
PT5   : 31/01/2005 JL V_60 FQ 11958 Ajout autre tests champ null
PT6   : 04/02/2005 JL V_60 FQ 11869 Ajout controle motif sortie DADS
PT7   : 04/02/2005 JL V_60 FQ 11549 Gestion edition des rib
PT8   : 20/07/2005 JL V_60 FQ 12442 Liste d'exportation supp pour edition fiche
                           salarié
PT9   : 15/11/2005 JL V_650 FQ 12612 Test convention collective non renseigné
PT10  : 22/09/2006 JL V_70 FQ 13503 Ajout test si condemploi=tps partiel et taux
                           tps partiel= 0 + correction 11/05/2007 + Modif 31/01/2007 : Code P au lieu de 002
PT11  : 27/12/2006 FC V_80 FQ 13313 Ajout des salariés sans date d'entrée et
                           sans établissement (issus DADS-U)
PT12  : 30/01/2007 FC V_80 Mise en place filtrage des habilitations/poupulations
PT13  : 07/06/2007 GGU V_72 FQ 14197 Prise en compte du filtre sur les
                           établissement et sur les dates d'entrée
PT14  : 07/06/2007 MF V_72 FQ 14247
PT15  : 04/07/2007 FC V_72 FQ 14515 Etat à 0 si des habilitations sont
                           paramétrées
PT16  : 13/07/2007 VG V_72 "Condition d''emploi" remplacé par "Caractéristique
                           activité" - FQ N°14568
PT17  : 24/07/2007 FC V_72 FQ 14423 Gestion des différents régimes
PT21  : 11/09/2008 SJ FQ 13798 Interdiction de saisie du point dans la commune de naissance
PT22  : 11/09/2008 SJ FQ 15138 ajouter le contrôle sur les caractères interdits sur la ville
PT23 : 25/09/2008 SJ FQ n°12376 Suppression du LanceEtatTob
PT24 : 29/09/2008 SJ FQ n°15103 Hergonomie sur l'onglet complement
}
Unit UTofPGEditFicheSalarie ;

Interface

Uses
     {$IFDEF VER150}
     Variants,
     {$ENDIF}
     StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}QRS1,EdtREtat,
{$ELSE}
     eQRS1,UtilEAGL,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,P5Def,EntPaie,ParamSoc,HQry,PGEditOutils,PgEditOutils2,ParamDat,PGOutils,PGoutils2,UTOB, PGDADSControles;

Type
  TOF_PGEDITFICHESAL = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
    Procedure OnClose                  ; override ;//PT23
    private
    TobEtat : Tob;
    procedure CkSortieClick (Sender : TObject);
    procedure DateElipsisClick (Sender : TObject);
    procedure ChoixRupture (Sender : TObject);
    procedure AccesRupture (Rupt : Integer;Cocher : Boolean);
    procedure ExitEdit (Sender : TObject);
    procedure EditionControleFicheSal ( Sender : TObject);
    procedure CreerFilleTobEtat(Etab,Salarie,Nom,Prenom,Theme,Zone,Commentaire : String) ;
    Function VerifCarApo (Donnee : String) : Boolean;
    Function VerifCarTU (Donnee : String) : Boolean;
    Function VerifCarNum (Donnee : String) : Boolean;
    Function VerifRepetApo (Donnee : String) : Boolean;
    Function VerifRepetTU (Donnee : String) : Boolean;
    Function VerifRepetEspace (Donnee : String) : Boolean;
    Function VerifCarPoint (Donnee : String) : Boolean; //PT21
  private
    Function GetStWhere : String;  //PT13
  end ;
var
Arg : String;

Implementation

procedure TOF_PGEDITFICHESAL.OnUpdate ;
var StWhere,ChampRupt,StOrder,StSQL : String;
    i : Integer;
//    Pages : TPageControl;
begin
  Inherited ;
//debut PT23
 If Arg = 'CONTROLE' then
    EditionControleFicheSal(Nil)
 else
 begin //fin PT23
  {Pages := TPageControl(GetControl('Pages'));
   StWhere := RecupWhereCritere(Pages);
   If GetCheckBoxState('CKSORTIE') = CbChecked then
   begin
        If StWhere <> ''  then StWhere := StWhere + 'AND (PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'" OR PSA_DATESORTIE>"'+UsDateTime(StrToDate(GetControlText('DATESORTIE')))+'")'  //PT3
        ELSE StWhere := 'WHERE (PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'" OR PSA_DATESORTIE>"'+UsDateTime(StrToDate(GetControlText('DATESORTIE')))+'")';
   end;}
   StWhere := GetStWhere; //PT13
   ChampRupt := '';
   For i := 1 to 6 do
   begin
        If GetCheckBoxState('CN'+IntToStr(i)) = CbChecked then
        begin
                If i<5 then ChampRupt := 'PSA_TRAVAILN'+IntToStr(i)
                else if i = 5 then ChampRupt := 'PSA_CODESTAT'
                else if i = 6 then ChampRupt := 'PSA_ETABLISSEMENT';
        end;
   end;
   SetControlText('XX_RUPTURE1',ChampRupt);
   If ChampRupt <> '' then StOrder := ' ORDER BY '+ChampRupt;
   If GetCheckBoxState('CALPHA') = CbChecked then
   begin
        If StOrder = '' then StOrder := ' ORDER BY PSA_LIBELLE'
        else StOrder := StOrder + ',PSA_LIBELLE';
   end
   else
   begin
        If StOrder = '' then StOrder := ' ORDER BY PSA_SALARIE'
        else StOrder := StOrder + ',PSA_SALARIE';
   end;
   If StWhere <> '' then StWhere := ' ' + StWhere;
   StSQL := 'SELECT * FROM SALARIES';
  //DEB PT12
  if Assigned(MonHabilitation) and (MonHabilitation.LeSQL <> '') then
    if StWhere <> '' then
      StWhere := StWhere + ' AND (' + MonHabilitation.LeSQL + ' OR PSA_ETABLISSEMENT = "")' //PT15
    else
      StWhere := '(' + MonHabilitation.LeSQL + ' OR PSA_ETABLISSEMENT = "")';               //PT15
  //FIN PT12
  TFQRS1(Ecran).WhereSQL := StSQL + StWHere + StOrder;
 end; //PT23
end ;

procedure TOF_PGEDITFICHESAL.OnArgument (S : String ) ;
var Num,i : Integer;
    Check : TCheckBox;
    Min,Max : String;
    Edit : THEdit;
begin
  Inherited ;
        Arg := ReadTokenPipe(S,';');
        For Num  :=  1 to VH_Paie.PGNbreStatOrg do
        begin
                if Num >4 then Break;
                VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)),GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));
              //  VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)+'_'),GetControl ('TA'+IntToStr(Num))); //PT24
                VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)+'_'),nil);//PT21
                SetControlVisible('CN'+IntToStr(Num),True);
                SetControlVisible('TA'+IntToStr(Num),true);    //PT21
        end;
        VisibiliteStat (GetControl ('PSA_CODESTAT'),GetControl ('TPSA_CODESTAT')) ;
        //VisibiliteStat (GetControl ('PSA_CODESTAT_'),GetControl ('TA5')) ; //PT21
        VisibiliteStat (GetControl ('PSA_CODESTAT_'),nil) ; //PT21
        SetControlVisible('TA5'+IntToStr(Num),true);
        if VH_Paie.PGLibCodeStat <> '' then setControlVisible('CN5',True);
        SetControlText('DOSSIER',GetParamSoc ('SO_LIBELLE'));
        SetControlText('DATEENTREE',DateToStr(IDate1900));
        SetControlText('DATEENTREE_',DateToStr(Date));
        Check := TCheckBox(GetControl('CKSORTIE'));
        If Check <> Nil Then Check.OnCLick := CkSortieClick;
        For i := 1 to 6 do
        begin
                Check := TCHeckBox(GetControl('CN'+IntToStr(i)));
                If Check <> Nil then Check.OnClick := ChoixRupture;
        end;
        RecupMinMaxTablette('PG','ETABLISS','ET_ETABLISSEMENT',Min,Max);
        SetControlText('ETABLISSEMENT',Min); //PT11
        SetControlText('ETABLISSEMENT_',Max); //PT11
        RecupMinMaxTablette('PG','SALARIES','PSA_SALARIE',Min,Max);
        SetControlText('PSA_SALARIE',Min);
        SetControlText('PSA_SALARIE_',Max);
        Edit := ThEdit(getcontrol('PSA_SALARIE'));
        If Edit <> nil then Edit.OnExit:=ExitEdit;
        Edit := ThEdit(getcontrol('PSA_SALARIE_'));
        If Edit <> nil then Edit.OnExit:=ExitEdit;
        Edit := THEdit(GetControl('DATEENTREE'));
        If Edit <> Nil Then Edit.OnElipsisClick := DateElipsisClick;
        Edit := THEdit(GetControl('DATEENTREE_'));
        If Edit <> Nil then Edit.OnElipsisClick := DateElipsisClick;
        Edit := THEdit(GetControl('DATESORTE'));
        If Edit <> Nil then Edit.OnElipsisClick := DateElipsisClick;
        If Arg = 'CONTROLE' then
        begin
                {TFQRS1(Ecran).BValider.OnClick := EditionControleFicheSal;} //PT23 mise en commentaire
                For i := 1 to 5 do SetControlVisible('CN'+IntToStr(i),False);
                TFQRS1(Ecran).Caption := 'Edition de contrôle des fiches salariés';
                UpdateCaption(Ecran);
                SetControlVisible('RIB',False);//PT7
        end
        else
        begin
             SetControlVisible('FLISTE',False);// PT8
             SetControlVisible('RIB',True);//PT7
        end;
end ;

procedure TOF_PGEDITFICHESAL.CkSortieClick(Sender : TObject);
begin
        If Sender = Nil then Exit;
        If GetCheckBoxState('CKSORTIE') = CbChecked then
        begin
                SetControlVisible('DATESORTIE',True);
                SetControlText('DATESORTIE',DateToStr(Date));
                SetControlVisible('TDATEARRET',True);
        end
        else
        begin
                SetControlVisible('DATESORTIE',False);
                SetControlVisible('TDATEARRET',False);
        end;
end;

procedure TOF_PGEDITFICHESAL.DateElipsisClick(Sender : TObject);
var key  :  char;
begin
    key  :=  '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGEDITFICHESAL.ChoixRupture (Sender : TObject);
var i : integer;
begin
        For i := 1 to 6 do
        begin
                If TCheckBox(Sender).Name = 'CN'+IntToStr(i) then
                begin
                        If GetCheckBoxState('CN'+IntToSTr(i)) = CbChecked then AccesRupture (i,True)
                        else AccesRupture(i,False);
                end;
        end;
end;

procedure TOF_PGEDITFICHESAL.AccesRupture (Rupt : Integer;Cocher : Boolean);
var i : Integer;
begin
        If Cocher = True then
        begin
                For i := 1 to 6 do
                begin
                        If i <> Rupt then
                        begin
                                SetControlChecked('CN'+IntToStr(i),False);
                                SetControlEnabled('CN'+IntToStr(i),False);
                        end;
                end;
        end
        else
        begin
                For i := 1 to 6 do SetControlEnabled('CN'+IntToStr(i),True);
        end;
end;
procedure TOF_PGEDITFICHESAL.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;

procedure TOF_PGEDITFICHESAL.EditionControleFicheSal ( Sender : TObject); //PT1
var
Q : TQuery;
TobSal : Tob;
Where : String;
Pages : TPageControl;
i,Num : Integer;
Sal,Nom,Prenom,Etabliss,StPages,NumSS,MotifSortie,MotifDads : String;
Communenaiss : String; //PT21
CP, Ville, Cinterdi, Adresse : string;
Resul : Integer;

//  Etab1,Etab2 : String; //PT11
begin
If TobEtat <> Nil Then FreeAndNil(TobEtat); //PT23

Pages:= TPageControl (GetControl ('Pages'));
{// Début PT11
Etab1:= GetcontrolText ('ETABLISSEMENT');
Etab2:= GetcontrolText ('ETABLISSEMENT_');

If Where <> ''  then
   Where:= Where+' AND ((PSA_ETABLISSEMENT>="'+Etab1+'" AND'+
           ' PSA_ETABLISSEMENT<="'+Etab2+'") OR PSA_ETABLISSEMENT="")'
else
   Where:= 'WHERE ((PSA_ETABLISSEMENT>="'+Etab1+'" AND'+
           ' PSA_ETABLISSEMENT<="'+Etab2+'") OR PSA_ETABLISSEMENT="")';

If Where <> ''  then
   Where:= Where+'AND ((PSA_DATEENTREE>="'+UsDateTime (StrToDate (GetControlText('DATEENTREE')))+'" AND'+
           ' PSA_DATEENTREE<="'+UsDateTime (StrToDate (GetControlText ('DATEENTREE_')))+'") OR'+
           ' (PSA_DATEENTREE IS NULL) OR'+
           ' (PSA_DATEENTREE="'+UsDateTime (IDate1900)+'"))'
else
   Where:= 'WHERE ((PSA_DATEENTREE>="'+UsDateTime (StrToDate (GetControlText('DATEENTREE')))+'" AND'+
           ' PSA_DATEENTREE<="'+UsDateTime (StrToDate (GetControlText ('DATEENTREE_')))+'") OR'+
           ' (PSA_DATEENTREE IS NULL) OR'+
           ' (PSA_DATEENTREE="'+UsDateTime (IDate1900)+'"))';
//Fin PT11 }
Where:= GetStWhere;//PT13

//DEB PT12
if Assigned(MonHabilitation) and (MonHabilitation.LeSQL <> '') then
   if Where <> '' then
      Where := Where + ' AND (' + MonHabilitation.LeSQL + ' OR PSA_ETABLISSEMENT = "")' //PT15
   else
      Where := '(' + MonHabilitation.LeSQL + ' OR PSA_ETABLISSEMENT = "")';               //PT15
//FIN PT12

Q:= OpenSQL ('SELECT PSA_ETABLISSEMENT, PSA_CONVENTION, PSA_SALARIE,'+
             ' PSA_LIBELLE, PSA_PRENOM, PSA_NUMEROSS, PSA_CODESTAT,'+
             ' PSA_TRAVAILN1, PSA_TRAVAILN2, PSA_TRAVAILN3, PSA_TRAVAILN4,'+
             ' PSA_LIBREPCMB1, PSA_LIBREPCMB2, PSA_LIBREPCMB3, PSA_LIBREPCMB4,'+
             ' PSA_NATIONALITE, PSA_DATENAISSANCE, PSA_DEPTNAISSANCE,'+
             ' PSA_PAYSNAISSANCE, PSA_REGIMESS, PSA_DADSCAT, PSA_DADSPROF,'+
             ' PSA_TYPEREGIME, PSA_REGIMEMAL, PSA_REGIMEAT, PSA_REGIMEVIP, PSA_REGIMEVIS,'+   //PT17
             ' PSA_DADSFRACTION, PSA_CONDEMPLOI, PSA_CODEEMPLOI,'+
             ' PSA_LIBELLEEMPLOI, PSA_MOTIFSORTIE, PSA_TAUXPARTIEL, PSA_DATEENTREE,'+    //PT11
             ' PSA_COMMUNENAISS, PSA_ADRESSE1, PSA_ADRESSE2, PSA_ADRESSE3, PSA_VILLE,'+
             ' PSA_CODEPOSTAL FROM SALARIES '+WHERE, True);
TobSal:= Tob.Create ('Les salariés', Nil, -1);
TobSal.LoadDetailDB ('Les salariés', '', '', Q, False);
Ferme (Q);
TobEtat:= Tob.Create ('Edition', Nil, -1);
For i := 0 to TobSal.Detail.Count -1 do
    begin
    Sal:= TobSal.Detail[i].GetValue ('PSA_SALARIE');
    If TobSal.Detail[i].GetValue ('PSA_LIBELLE')<>null then
       Nom:= TobSal.Detail[i].GetValue ('PSA_LIBELLE')
    else
       Nom:= '';

    if TobSal.Detail[i].GetValue ('PSA_PRENOM')<>null then
       Prenom:= TobSal.Detail[i].GetValue ('PSA_PRENOM')
    else
       Prenom:= '';

    Etabliss:= TobSal.Detail[i].GetValue ('PSA_ETABLISSEMENT');
    //DEB PT11
    if Etabliss = '' then
      CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Identité', 'Etablissement',
                             'Zone non renseignée');
    //FIN PT11

    If TobSal.Detail[i].GetValue ('PSA_NUMEROSS')<>null then
       NumSS:= TobSal.Detail[i].GetValue ('PSA_NUMEROSS')
    else
       NumSS:= '';

    If Nom<>'' then
       begin
       If Not VerifCarApo (Nom) then
          CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Identité', 'Nom',
                             'Ne pas mettre d''apostrophe en début ou fin de zone');
       If Not VerifCarTU (Nom) then
          CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Identité', 'Nom',
                             'Ne pas mettre de trait d''union en début ou fin de zone');
       If Not VerifCarNum (Nom) then
          CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Identité', 'Nom',
                             'Ne pas mettre de caractère numérique');
       If Not VerifRepetApo (Nom) then
          CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Identité', 'Nom',
                             'Ne pas mettre des répétitions d''apostrophe');
       If Not VerifRepetTU (Nom) then
          CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Identité', 'Nom',
                             'Ne pas mettre des répétitions de trait d''union');
       If Not VerifRepetEspace (Nom) then
          CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Identité', 'Nom',
                             'Ne pas mettre des répétitions d''espace');
       end
    else
       CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Identité', 'Nom',
                          'Zone non renseignée');

    If Prenom<>'' then
       begin
       If Not VerifCarApo (Prenom) then
          CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Identité', 'Prénom',
                             'Ne pas mettre d''apostrophe en début ou fin de zone');
       If Not VerifCarTU (Prenom) then
          CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Identité', 'Prénom',
                             'Ne pas mettre de trait d''union en début ou fin de zone');
       If Not VerifCarNum (Prenom) then
          CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Identité', 'Prénom',
                             'Ne pas mettre de caractère numérique');
       If Not VerifRepetApo (Prenom) then
          CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Identité', 'Prénom',
                             'Ne pas mettre des répétitions d''apostrophe');
       If Not VerifRepetTU (Prenom) then
          CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Identité', 'Prénom',
                             'Ne pas mettre des répétitions de trait d''union');
       If Not VerifRepetEspace (Prenom) then
          CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Identité', 'Prénom',
                             'Ne pas mettre des répétitions d''espace');
       end
    else
       CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Identité', 'Prénom',
                          'Zone non renseignée');
//deb PT22

    if TobSal.Detail[i].GetValue ('PSA_ADRESSE1')<>null then
       Adresse:= TobSal.Detail[i].GetValue ('PSA_ADRESSE1')
    else
       Adresse:= '';
    If Adresse <>'' then
    Begin
    	Resul := ControleCar(Adresse , '1', false, Cinterdi);
   	  If Resul<>0 then
	      CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Identité', 'Adresse 1',
                             'Caractère '+'"'+Cinterdi+'"'+' interdit');
    End ;

    if TobSal.Detail[i].GetValue ('PSA_ADRESSE2')<>null then
       Adresse:= TobSal.Detail[i].GetValue ('PSA_ADRESSE2')
    else
       Adresse:= '';
    If Adresse <>'' then
    Begin
    	Resul := ControleCar(Adresse , '1', false, Cinterdi);
   	  If Resul<>0 then
	      CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Identité', 'Adresse 2',
                             'Caractère '+'"'+Cinterdi+'"'+' interdit');
 	  End;

    if TobSal.Detail[i].GetValue ('PSA_ADRESSE3')<>null then
       Adresse:= TobSal.Detail[i].GetValue ('PSA_ADRESSE3')
    else
       Adresse:= '';
    If Adresse <>'' then
    Begin
    	Resul := ControleCar(Adresse , '1', false, Cinterdi);
   	  If Resul<>0 then
	      CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Identité', 'Adresse 3',
                             'Caractère '+'"'+Cinterdi+'"'+' interdit');
 	  End;

    if TobSal.Detail[i].GetValue ('PSA_VILLE')<>null then
       Ville:= TobSal.Detail[i].GetValue ('PSA_VILLE')
    else
       Ville:= '';
    If Ville <>'' then
    Begin
    	Resul := ControleCar(Ville , '7', false, Cinterdi);
   	  If Resul<>0 then
	      CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Identité', 'Ville',
                             'Caractère '+'"'+Cinterdi+'"'+' interdit');
 	  End;

    if TobSal.Detail[i].GetValue ('PSA_CODEPOSTAL')<>null then
       CP:= TobSal.Detail[i].GetValue ('PSA_CODEPOSTAL')
    else
       CP:= '';
    If CP <>'' then
    Begin
    	Resul := ControleCar(CP , '6', false, Cinterdi);
   	  If Resul<>0 then
	      CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Identité', 'CODE POSTAL',
                             'Caractère '+'"'+Cinterdi+'"'+' interdit');
 	  End;

    
//fin PT22
{PT14
    If NumSS='' then
       CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Identité',
                          'N° sécurité sociale', 'Zone non renseignée')
}
    If NumSS='' then
       CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Etat civil',
                          'N° sécurité sociale', 'Zone non renseignée')
    else
       begin
{PT14
       If (copy (NumSS, 1, 1)='7') or (copy (NumSS, 1, 1)='8') then
          CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Identité',
                             'N° sécurité sociale', 'N° provisoire');
}
       If (copy (NumSS, 1, 1)='7') or (copy (NumSS, 1, 1)='8') then
          CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Etat civil',
                             'N° sécurité sociale', 'N° provisoire');
       end;

   If TobSal.Detail[i].GetValue ('PSA_CONVENTION')='' then
      CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Affectation',
                         'Convention collective', 'Zone non renseignée');

   if VH_Paie.PGLibCodeStat<>'' then
      begin
      if TobSal.Detail[i].GetValue ('PSA_CODESTAT')<>Null then
         begin
         If TobSal.Detail[i].GetValue ('PSA_CODESTAT')='' then
            CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Affectation',
                               'Code statistique', 'Zone non renseignée');
         end
      else
         CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Affectation',
                            'Code statistique', 'Zone non renseignée');
      end;

   For num := 1 to VH_Paie.PGNbreStatOrg do
       begin
       If TobSal.Detail[i].GetValue ('PSA_TRAVAILN'+IntToStr (Num))<>Null then
          begin
          If TobSal.Detail[i].GetValue ('PSA_TRAVAILN'+IntToStr (Num))='' then
             begin
             If Num=1 then
                CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Affectation',
                                   VH_Paie.PGLibelleOrgStat1,
                                   'Zone non renseignée')
             else
             If Num=2 then
                CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Affectation',
                                   VH_Paie.PGLibelleOrgStat2,
                                   'Zone non renseignée')
             else
             If Num=3 then
                CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Affectation',
                                   VH_Paie.PGLibelleOrgStat3,
                                   'Zone non renseignée')
             else
             If Num=4 then
                CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Affectation',
                                   VH_Paie.PGLibelleOrgStat4,
                                   'Zone non renseignée');
             end;
          end
       else
          begin
          If Num=1 then
             CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Affectation',
                                VH_Paie.PGLibelleOrgStat1,
                                'Zone non renseignée')
          else
          If Num=2 then
             CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Affectation',
                                VH_Paie.PGLibelleOrgStat2,
                                'Zone non renseignée')
          else
          If Num=3 then
             CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Affectation',
                                VH_Paie.PGLibelleOrgStat3,
                                'Zone non renseignée')
          else
          If Num=4 then
             CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Affectation',
                                VH_Paie.PGLibelleOrgStat4,
                                'Zone non renseignée');
          end;
       end;

   For num := 1 to VH_Paie.PgNbCombo do
       begin
       if TobSal.Detail[i].GetValue ('PSA_LIBREPCMB'+IntToStr(Num))<>Null then
          begin
          If TobSal.Detail[i].GetValue ('PSA_LIBREPCMB'+IntToStr(Num))='' then
             begin
             If Num=1 then
                CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Zones libres',
                                   VH_Paie.PgLibCombo1, 'Zone non renseignée')
             else
             If Num=2 then
                CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Zones libres',
                                   VH_Paie.PgLibCombo2, 'Zone non renseignée')
             else
             If Num=3 then
                CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Zones libres',
                                   VH_Paie.PgLibCombo3, 'Zone non renseignée')
             else
             If Num=4 then
                CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Zones libres',
                                   VH_Paie.PgLibCombo4, 'Zone non renseignée');
             end;
          end
       else
          begin
          If Num=1 then
             CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Zones libres',
                                VH_Paie.PgLibCombo1, 'Zone non renseignée')
          else
          If Num=2 then
             CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Zones libres',
                                VH_Paie.PgLibCombo2, 'Zone non renseignée')
          else
          If Num=3 then
             CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Zones libres',
                                VH_Paie.PgLibCombo3, 'Zone non renseignée')
          else
          If Num=4 then
             CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Zones libres',
                                VH_Paie.PgLibCombo4, 'Zone non renseignée');
          end;
       end;

   If TobSal.Detail[i].GetValue ('PSA_NATIONALITE')<>null then
      begin
      if TobSal.Detail[i].GetValue ('PSA_NATIONALITE')='' then
         CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Etat civil',
                            'Nationalité', 'Zone non renseignée');
      end
   else
      CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Etat civil',
                         'Nationalité', 'Zone non renseignée');

   If TobSal.Detail[i].GetValue ('PSA_DATENAISSANCE')<>Null then
      begin
      If TobSal.Detail[i].GetValue ('PSA_DATENAISSANCE')=IDate1900 then
         CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Etat civil',
                            'Date de naissance', 'Zone non renseignée');
      end
   else
      CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Etat civil',
                         'Date de naissance', 'Zone non renseignée');

   If TobSal.Detail[i].GetValue ('PSA_DEPTNAISSANCE')<>Null then
      begin
      If TobSal.Detail[i].GetValue ('PSA_DEPTNAISSANCE')='' then
         CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Etat civil',
                            'Département de naissance', 'Zone non renseignée');
      end
   else
      CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Etat civil',
                         'Département de naissance', 'Zone non renseignée');

   If TobSal.Detail[i].GetValue ('PSA_PAYSNAISSANCE')<>Null then
      begin
      If TobSal.Detail[i].GetValue ('PSA_PAYSNAISSANCE')='' then
         CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Etat civil',
                            'Pays de naissance', 'Zone non renseignée');
      end
   else
      CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Etat civil',
                         'Pays de naissance', 'Zone non renseignée');

// deb PT21
    if TobSal.Detail[i].GetValue ('PSA_COMMUNENAISS')<>null then
       Communenaiss:= TobSal.Detail[i].GetValue ('PSA_COMMUNENAISS')
    else
       Communenaiss:= '';

    If Communenaiss<>'' then
    begin
       If Not VerifCarPoint (Communenaiss) then
          CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Etat civil', 'Commune de naissance',
                             'Caractère "point" interdit');
    end;
//fin PT21
  //DEB PT11
   If TobSal.Detail[i].GetValue ('PSA_DATEENTREE')<>Null then
      begin
      If TobSal.Detail[i].GetValue ('PSA_DATEENTREE')=IDate1900 then
         CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Emploi',
                            'Date d''entrée', 'Zone non renseignée');
      end
   else
      CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'PSA_DATEENTREE',
                         'Date d''entrée', 'Zone non renseignée');
   //FIN PT11

   //DEB PT17
   If TobSal.Detail[i].GetValue ('PSA_TYPEREGIME')<>Null then
   begin
     If TobSal.Detail[i].GetValue ('PSA_TYPEREGIME')='-' then
     begin
       If TobSal.Detail[i].GetValue ('PSA_REGIMESS')<>Null then
          begin
          If TobSal.Detail[i].GetValue ('PSA_REGIMESS')='' then
             CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'DADS/Fiscal',
                                'Régime SS', 'Zone non renseignée');
          end
       else
          CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'DADS/Fiscal', 'Régime SS',
                             'Zone non renseignée');
     end
     else
     begin
       If TobSal.Detail[i].GetValue ('PSA_REGIMEMAL')<>Null then
          begin
          If TobSal.Detail[i].GetValue ('PSA_REGIMEMAL')='' then
             CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'DADS/Fiscal',
                                'Régime obligatoire risque maladie', 'Zone non renseignée');
          end
       else
          CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'DADS/Fiscal', 'Régime obligatoire risque maladie',
                             'Zone non renseignée');
       If TobSal.Detail[i].GetValue ('PSA_REGIMEAT')<>Null then
          begin
          If TobSal.Detail[i].GetValue ('PSA_REGIMEAT')='' then
             CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'DADS/Fiscal',
                                'Régime obligatoire risque AT', 'Zone non renseignée');
          end
       else
          CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'DADS/Fiscal', 'Régime obligatoire risque AT',
                             'Zone non renseignée');
       If TobSal.Detail[i].GetValue ('PSA_REGIMEVIP')<>Null then
          begin
          If TobSal.Detail[i].GetValue ('PSA_REGIMEVIP')='' then
             CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'DADS/Fiscal',
                                'Régime obligatoire vieillesse (PP)', 'Zone non renseignée');
          end
       else
          CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'DADS/Fiscal', 'Régime obligatoire vieillesse (PP)',
                             'Zone non renseignée');
       If TobSal.Detail[i].GetValue ('PSA_REGIMEVIS')<>Null then
          begin
          If TobSal.Detail[i].GetValue ('PSA_REGIMEVIS')='' then
             CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'DADS/Fiscal',
                                'Régime obligatoire vieillesse (PS)', 'Zone non renseignée');
          end
       else
          CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'DADS/Fiscal', 'Régime obligatoire vieillesse (PS)',
                             'Zone non renseignée');
     end;
   end
   else
   begin
     If TobSal.Detail[i].GetValue ('PSA_REGIMESS')<>Null then
        begin
        If TobSal.Detail[i].GetValue ('PSA_REGIMESS')='' then
           CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'DADS/Fiscal',
                              'Régime SS', 'Zone non renseignée');
        end
     else
        CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'DADS/Fiscal', 'Régime SS',
                           'Zone non renseignée');
   end;
   //FIN PT17

   If TobSal.Detail[i].GetValue ('PSA_DADSCAT')<>Null then
      begin
      If TobSal.Detail[i].GetValue ('PSA_DADSCAT')='' then
         CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'DADS/Fiscal',
                            'Statut catégoriel', 'Zone non renseignée');
      end
   else
      CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'DADS/Fiscal',
                         'Statut catégoriel', 'Zone non renseignée');

   If TobSal.Detail[i].GetValue ('PSA_DADSPROF')<>Null then
      begin
      If TobSal.Detail[i].GetValue ('PSA_DADSPROF')='' then
         CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'DADS/Fiscal',
                            'Statut professionnel', 'Zone non renseignée');
      end
   else
      CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'DADS/Fiscal',
                         'Statut professionnel', 'Zone non renseignée');

   If TobSal.Detail[i].GetValue ('PSA_DADSFRACTION')<>Null then
      begin
      If TobSal.Detail[i].GetValue ('PSA_DADSFRACTION')='' then
         CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'DADS/Fiscal',
                            'Fraction DADS', 'Zone non renseignée');
      end
   else
      CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'DADS/Fiscal',
                         'Fraction DADS', 'Zone non renseignée');

   If TobSal.Detail[i].GetValue ('PSA_CONDEMPLOI')<>Null then
      begin
      If TobSal.Detail[i].GetValue ('PSA_CONDEMPLOI')='' then
{PT16
         CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'DADS/Fiscal',
                            'Condition emploi', 'Zone non renseignée')
}
         CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'DADS/Fiscal',
                            'Caractéristique activité', 'Zone non renseignée')
//FIN PT16
//DEBUT PT10
      else
      If TobSal.Detail[i].GetValue ('PSA_CONDEMPLOI')='P' then
         begin
         If TobSal.Detail[i].GetValue ('PSA_TAUXPARTIEL')<>Null then
            begin
            If TobSal.Detail[i].GetValue ('PSA_TAUXPARTIEL')<=0 then
               CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'DADS/Fiscal',
                                  'Taux temps partiel', 'Zone non renseignée')
            end
         else
            CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'DADS/Fiscal',
                               'Taux temps partiel', 'Zone non renseignée');
         end;
//FIN PT10
      end
   else
{PT16
      CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'DADS/Fiscal',
                         'Condition emploi', 'Zone non renseignée');
}
      CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'DADS/Fiscal',
                         'Caractéristique activité', 'Zone non renseignée');
//FIN PT16

   If TobSal.Detail[i].GetValue ('PSA_CODEEMPLOI')<>Null then
      begin
      If TobSal.Detail[i].GetValue ('PSA_CODEEMPLOI')='' then
         CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Affectation',
                            'Nomenclature PCS', 'Zone non renseignée');
      end
   else
      CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Affectation',
                         'Nomenclature PCS', 'Zone non renseignée');

   If TobSal.Detail[i].GetValue ('PSA_LIBELLEEMPLOI')<>Null then
      begin
      If TobSal.Detail[i].GetValue ('PSA_LIBELLEEMPLOI')='' then
         CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Affectation',
                            'Libellé emploi', 'Zone non renseignée');
      end
   else
      CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Affectation',
                         'Libellé emploi', 'Zone non renseignée');

   if TobSal.Detail[i].GetValue ('PSA_MOTIFSORTIE')<>null then
      MotifSortie:= TobSal.Detail[i].GetValue ('PSA_MOTIFSORTIE')
   else
      MotifSortie:= '';

   If MotifSortie<>'' then
      begin
      Q:= OpenSQL ('SELECT PMS_DADSU'+
                   ' FROM MOTIFSORTIEPAY WHERE'+
                   ' PMS_CODE="'+MotifSortie+'"',True);
      If Not Q.Eof then
         MotifDads:= Q.FindField ('PMS_DADSU').AsString
      else
         MotifDads:= '';
      Ferme (Q);
      If MotifDads='000' then
         CreerFilleTobEtat (Etabliss, Sal, Nom, Prenom, 'Emploi',
                            'Motif de sortie',
                            'Le code DADSU correspondant à ce motif est 000');
      end;
   end;

If GetCheckBoxState ('CN6')=CbChecked then
   begin
   if GetCheckBoxState ('CALPHA')=CbChecked then
      TobEtat.Detail.Sort ('PSA_ETABLISSEMENT;PSA_LIBELLE;PSA_SALARIE;THEME')
   else
      TobEtat.Detail.Sort ('PSA_ETABLISSEMENT;PSA_SALARIE;THEME');
   end
else
   begin
   if GetCheckBoxState ('CALPHA')=CbChecked then
      TobEtat.Detail.Sort ('PSA_LIBELLE;PSA_SALARIE;THEME')
   else
      TobEtat.Detail.Sort ('PSA_SALARIE;THEME');
   end;

TobSal.Free;
//debut PT23
TFQRS1(Ecran).LaTob:= TobEtat;
TFQRS1(Ecran).NatureEtat := 'PSA';
TFQRS1(Ecran).CodeEtat := 'PCS';
{StPages:= '';
{$IFDEF EAGLCLIENT}
{StPages:= AglGetCriteres (Pages, FALSE);
{$ENDIF}
{If GetCheckBoxState ('FLISTE')=CbChecked then
   LanceEtatTOB ('E', 'PSA', 'PCS', TobEtat, True, True, False, Pages, ' ', '',
                 False, 0, StPages)
else
   LanceEtatTOB ('E', 'PSA', 'PCS', TobEtat, True, False, False, Pages, ' ', '',
                 False, 0, StPages);
If TobEtat<>Nil then
 TobEtat.Free;}
// fin PT23
end;

procedure TOF_PGEDITFICHESAL.CreerFilleTobEtat(Etab,Salarie,Nom,Prenom,Theme,Zone,Commentaire : String) ;
var T : TOB;
begin
        T := Tob.Create ('la tob',TobEtat,-1);
        T.AddChampSupValeur('PSA_ETABLISSEMENT',Etab,False);
        T.AddChampSupValeur('PSA_SALARIE',Salarie,False);
        T.AddChampSupValeur('PSA_LIBELLE',Nom,False);
        T.AddChampSupValeur('PSA_PRENOM',Prenom,False);
        T.AddChampSupValeur('THEME',Theme,False);
        T.AddChampSupValeur('ZONE',Zone,False);
        T.AddChampSupValeur('COMMENTAIRE',Commentaire,False);
end;

Function TOF_PGEDITFICHESAL.VerifCarApo (Donnee : String) : Boolean;
var L : Integer;
begin
        L := Length(Donnee);
        result := True;
        If L >= 1 then
        begin
                If Donnee[1] = '''' then Result := False;
                If Donnee[L] = '''' then Result := False;
        end;
end;

Function TOF_PGEDITFICHESAL.VerifCarTU (Donnee : String) : Boolean;
var L : Integer;
begin
        L := Length(Donnee);
        result := True;
        If L >= 1 then
        begin
                If Donnee[1] = '-' then Result := False;
                If Donnee[L] = '-' then Result := False;
        end;
end;

Function TOF_PGEDITFICHESAL.VerifCarNum (Donnee : String) : Boolean;
var L,i : Integer;
    Ch : CHar;
begin
        L := Length(Donnee);
        result := True;
        For i := 1 to L do
        begin
                Ch := Donnee[i];
                If (Ch in [',','0','1','2','3','4','5','6','7','8','9']) then Result := False;
        end;
end;

Function TOF_PGEDITFICHESAL.VerifRepetApo (Donnee : String) : Boolean;
var L,i : Integer;
begin
        L := Length(Donnee);
        result := True;
        For i := 1 to L - 1 do
        begin
                If (Donnee[i] = '''') and (Donnee[i+1] = '''') then Result := False;
        end;
end;
Function TOF_PGEDITFICHESAL.VerifRepetTU (Donnee : String) : Boolean;
var L,i : Integer;
begin
        L := Length(Donnee);
        result := True;
        For i := 1 to L - 1 do
        begin
                If (Donnee[i] = '-') and (Donnee[i+1] = '-') then Result := False;
        end;
end;
Function TOF_PGEDITFICHESAL.VerifRepetEspace (Donnee : String) : Boolean;
var L,i : Integer;
begin
        L := Length(Donnee);
        result := True;
        For i := 1 to L - 1 do
        begin
                If (Donnee[i] = ' ') and (Donnee[i+1] = ' ') then Result := False;
        end;
end;

//deb PT21
Function  TOF_PGEDITFICHESAL.VerifCarPoint(Donnee : String) : Boolean;
var L,i : Integer;
    Ch : CHar;
begin
        L := Length(Donnee);
        result := True;
        For i := 1 to L do
        begin
                Ch := Donnee[i];
                If Ch = '.' then
                 Result := False;
        end;
end;
//fin PT21

function TOF_PGEDITFICHESAL.GetStWhere: String;
begin
  result := RecupWhereCritere(TPageControl(GetControl('Pages')));
  //DEB PT15
  If result <> ''  then
  begin
    if (GetcontrolText('ETABLISSEMENT') <> '') and (GetcontrolText('ETABLISSEMENT_') <> '') then
      result := result + ' AND ((PSA_ETABLISSEMENT>="' + GetcontrolText('ETABLISSEMENT') + '" AND PSA_ETABLISSEMENT<="'+GetcontrolText('ETABLISSEMENT_')+'") OR PSA_ETABLISSEMENT="")'
    else if (GetcontrolText('ETABLISSEMENT') = '') and (GetcontrolText('ETABLISSEMENT_') <> '') then
      result := result + ' AND ((PSA_ETABLISSEMENT<="'+GetcontrolText('ETABLISSEMENT_')+'") OR PSA_ETABLISSEMENT="")'
    else if (GetcontrolText('ETABLISSEMENT') <> '') and (GetcontrolText('ETABLISSEMENT_') = '') then
      result := result + ' AND ((PSA_ETABLISSEMENT>="'+GetcontrolText('ETABLISSEMENT')+'") OR PSA_ETABLISSEMENT="")';
  end
  else
  begin
    if (GetcontrolText('ETABLISSEMENT') <> '') and (GetcontrolText('ETABLISSEMENT_') <> '') then
      result := ' WHERE ((PSA_ETABLISSEMENT>="' + GetcontrolText('ETABLISSEMENT') + '" AND PSA_ETABLISSEMENT<="'+GetcontrolText('ETABLISSEMENT_')+'") OR PSA_ETABLISSEMENT="")'
    else if (GetcontrolText('ETABLISSEMENT') = '') and (GetcontrolText('ETABLISSEMENT_') <> '') then
      result := ' WHERE ((PSA_ETABLISSEMENT<="'+GetcontrolText('ETABLISSEMENT_')+'") OR PSA_ETABLISSEMENT="")'
    else if (GetcontrolText('ETABLISSEMENT') <> '') and (GetcontrolText('ETABLISSEMENT_') = '') then
      result := ' WHERE ((PSA_ETABLISSEMENT>="'+GetcontrolText('ETABLISSEMENT')+'") OR PSA_ETABLISSEMENT="")';
  end;
  //FIN PT15
  If result <> ''  then
    result := result + 'AND ((PSA_DATEENTREE>="'+UsDateTime(StrToDate(GetControlText('DATEENTREE')))+'" AND PSA_DATEENTREE<="'+UsDateTime(StrToDate(GetControlText('DATEENTREE_')))+'") OR (PSA_DATEENTREE IS NULL) OR (PSA_DATEENTREE="'+UsDateTime(IDate1900)+'"))'
  else
    result := 'WHERE ((PSA_DATEENTREE>="'+UsDateTime(StrToDate(GetControlText('DATEENTREE')))+'" AND PSA_DATEENTREE<="'+UsDateTime(StrToDate(GetControlText('DATEENTREE_')))+'") OR (PSA_DATEENTREE IS NULL) OR (PSA_DATEENTREE="'+UsDateTime(IDate1900)+'"))';
   If GetCheckBoxState('CKSORTIE') = CbChecked then
   begin
     If result <> ''  then result := result + 'AND (PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'" OR PSA_DATESORTIE>"'+UsDateTime(StrToDate(GetControlText('DATESORTIE')))+'")'  //PT3
     ELSE result := 'WHERE (PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'" OR PSA_DATESORTIE>"'+UsDateTime(StrToDate(GetControlText('DATESORTIE')))+'")';
   end;
end;
//debut PT23
procedure TOF_PGEDITFICHESAL.OnClose;
Begin
	If TobEtat <> Nil Then FreeAndNil(TobEtat);
End;
//fin PT23
Initialization
  registerclasses ( [ TOF_PGEDITFICHESAL ] ) ;
end.

