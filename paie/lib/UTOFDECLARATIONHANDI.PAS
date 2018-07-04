{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 31/05/2001
Modifié le ... : 05/10/2001
Description .. : Source TOF edition de la déclaration des travailleurs 
Suite ........ : handicapés
Mots clefs ... : PAIE;HANDICAPE          
*****************************************************************
PT- 1 : 10/10/2001 : JL 562: Procédures de vérification des dates :verifdateedition et verifdateagrement + calendrier
PT- 2 : 10/10/2001 : JL 562 : Ajout de l'affichage de l'effectif par défaut dans la partie 3c
PT- 3 : 11/10/2001 : JL 562 : Vérification des codes PCS
PT- 4 : 09/01/2002 : JL 571 : Contrôle nombre salariés au 31/12 < 20
PT- 5 : 03/04/2002 : JL 571 : Fiche Bug N° 395 corection faute d'orthographe
PT- 6 : 06/06/2002 : JL 582 : Fiche de bug n°10150 Gestion du cas ou un seul exercice social.
PT- 7 ! 09/10/2003 : JL V_42 Gestion changement code PCS
PT- 8 ! 23/02/2004 : JL V_50 Gestion nouveau code emploi spécifiques
PT- 9 : 31/01/2005 : JL V_60 Modifs calculs effectifs
---- JL 20/03/2006 modification clé annuaire ----
PT10  : 04/04/2006 : JL V_65 FQ 12946 Gestion établissement de rattachement
PT11  : 07/02/2007 : FC V_80 Mise en place filtrage habilitations/populations
PT12  : 09/08/2007 : JL V_80 Saisie établissement obligatoire
}
Unit UTOFDECLARATIONHANDI ;

Interface

Uses
     {$IFDEF VER150}
     Variants,
     {$ENDIF}
     StdCtrls, Controls, Classes,forms, sysutils,ComCtrls,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,QRS1,DBCtrls,
{$ELSE}
     eQrs1,UtileAGL,
{$ENDIF}
     HCtrls, HEnt1, HMsgBox, UTOF,ParamDat,UTOB,PGEditOutils,EntPaie,HTB97,
     P5Def //PT11
     ;

Type
  TOF_DECLARATIONHANDI = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    private
    Habilitation:String;  //PT11
    procedure effectifs(Sender : Tobject);
    procedure partie3a(Sender : Tobject);
    procedure partie3b(Sender : Tobject);
    procedure partie3c(Sender : Tobject);
    procedure partie3d(Sender : Tobject);
    procedure ChangePCS(Sender: Tobject);
    procedure ChangeT2(Sender: Tobject);
    procedure ChangeT1(Sender: Tobject);
    procedure ChangeAdresse(Sender: Tobject);
    procedure ChangeAccord(Sender: Tobject);
    procedure AccesCondition;
    procedure verifdateedition(Sender: Tobject);
    procedure verifdateagrement(Sender: Tobject);
    procedure DateElipsisclick(Sender: TObject);
    procedure VerifPCS(Sender: Tobject);
    procedure VerifEffectif(Sender:TObject);  //PT 4
  end ;

Implementation

procedure TOF_DECLARATIONHANDI.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_DECLARATIONHANDI.OnDelete ;
begin
  Inherited ;
end ;
// Confection de la requete en focntion des critères de la fiche QRS1
procedure TOF_DECLARATIONHANDI.OnUpdate ;
Var  requete,Etab:String;
     Etablissement: THValComboBox ;
     Assiette:integer;
     Resultat:Double;
     Total:Real;
     Bt : TToolBarButton97;
begin
     Inherited ;
     //DEBUT PT12
     Bt := TToolBarButton97(GetControl('BValider'));
     If GetControlText('XX_VARETAB') = '' then
     begin
      Bt.ModalResult := -1;
      PGIBox('Vous devez renseigner l''établissement',Ecran.Caption);
      SetFocusControl('XX_VARETAB');
     end
     else Bt.ModalResult := 0;
     //FIN PT12
     Etablissement:=THValComboBox(GetControl('XX_VARETAB'));
     Etab:=Etablissement.value;
     Assiette:=StrToInt(GetControlText('XX_VART1T2'));
     Resultat:=Assiette * 6 / 100;
     Total:=int(Resultat);
     SetControlText('XX_VARTOTAL',FloatToStr(Total));
        requete:='SELECT PSA_SEXE,PSA_CODEEMPLOI,PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_DATENAISSANCE,PSA_DATEENTREE,'+
                 ' PSA_UNITEPRISEFF,PSA_PRISEFFECTIF,'+
                 ' PGH_SALARIE,PGH_COTOREP,PGH_CATEGHANDI,PGH_PLACEANT,PGH_ACCTRAV,PGH_TXINV,PGH_INVPEN,PGH_FORMATION'+
                 ' ,PGH_MUTILE,PGH_ASSMUTILE,PGH_HANDICAP,'+
                 ' ET_ETABLISSEMENT,ET_SIRET,ET_APE,ET_LIBELLE,ET_ADRESSE1,ET_ADRESSE2,ET_ADRESSE3,ET_CODEPOSTAL,ET_VILLE,'+
                 ' PCI_FINCONTRAT,PCI_TYPECONTRAT,PCI_DEBUTCONTRAT'+
                 ' FROM HANDICAPE'+
                 ' LEFT JOIN SALARIES ON PSA_SALARIE=PGH_SALARIE'+
                 ' LEFT JOIN ETABLISS ON ET_ETABLISSEMENT=PSA_ETABLISSEMENT'+
                 ' LEFT JOIN ETABCOMPL ON ETB_ETABLISSEMENT=ET_ETABLISSEMENT'+  //PT10
                 ' LEFT JOIN CONTRATTRAVAIL C1 ON PCI_SALARIE=PSA_SALARIE'+
                 ' WHERE (ET_ETABLISSEMENT="'+Etab+'" OR ETB_RATTACHEHANDI="'+Etab+'")'+ //PT10
                 ' AND (PGH_HANDICAP="X" OR PGH_COTOREP="X")'+
                 ' AND PSA_PRISEFFECTIF="X"'+ Habilitation + //PT11
                 ' AND PCI_DEBUTCONTRAT IN'+
                 ' (SELECT MAX(PCI_DEBUTCONTRAT) FROM CONTRATTRAVAIL C2 WHERE C1.PCI_SALARIE=C2.PCI_SALARIE GROUP BY PCI_salarie)'+
                 ' ORDER BY ET_ETABLISSEMENT,PGH_SALARIE';
     If existeSQL('SELECT PGH_SALARIE FROM HANDICAPE LEFT JOIN SALARIES ON PGH_SALARIE=PSA_SALARIE'+
                  ' LEFT JOIN ETABLISS ON ET_ETABLISSEMENT=PSA_ETABLISSEMENT'+
                  ' LEFT JOIN ETABCOMPL ON ETB_ETABLISSEMENT=ET_ETABLISSEMENT'+  //PT10
                  ' LEFT JOIN CONTRATTRAVAIL C1 ON PCI_SALARIE=PSA_SALARIE'+
                  ' WHERE (PGH_HANDICAP="X" OR PGH_COTOREP="X") AND PSA_PRISEFFECTIF="X"'+ Habilitation + //PT11
                  ' AND (ET_ETABLISSEMENT="'+Etab+'" OR ETB_RATTACHEHANDI="'+Etab+'") AND PCI_DEBUTCONTRAT IN'+
                 ' (SELECT MAX(PCI_DEBUTCONTRAT) FROM CONTRATTRAVAIL C2 WHERE C1.PCI_SALARIE=C2.PCI_SALARIE GROUP BY PCI_salarie)') Then
     begin
     SetControlChecked('XX_VARSALARIE',True);
     TFQRS1(Ecran).WhereSQL:=requete;
     end
     Else
     Begin
     TFQRS1(Ecran).WhereSQL:='SELECT PSA_SEXE,PSA_CODEEMPLOI,PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_DATENAISSANCE,PSA_DATEENTREE,'+
                             ' PSA_UNITEPRISEFF,PSA_PRISEFFECTIF,'+
                             ' PGH_SALARIE,PGH_COTOREP,PGH_CATEGHANDI,PGH_PLACEANT,PGH_ACCTRAV,PGH_TXINV,PGH_INVPEN,PGH_FORMATION'+
                             ' ,PGH_MUTILE,PGH_ASSMUTILE,PGH_HANDICAP,'+
                             ' ET_ETABLISSEMENT,ET_SIRET,ET_APE,ET_LIBELLE,ET_ADRESSE1,ET_ADRESSE2,ET_ADRESSE3,ET_CODEPOSTAL,ET_VILLE,'+
                             ' PCI_FINCONTRAT,PCI_TYPECONTRAT,PCI_DEBUTCONTRAT'+
                             ' FROM SALARIES'+
                             ' LEFT JOIN ETABLISS ON PSA_ETABLISSEMENT=ET_ETABLISSEMENT'+
                             ' LEFT JOIN ETABCOMPL ON ETB_ETABLISSEMENT=ET_ETABLISSEMENT'+  //PT10
                             ' LEFT JOIN HANDICAPE ON PSA_SALARIE=PGH_SALARIE'+
                             ' LEFT JOIN CONTRATTRAVAIL C1 ON PCI_SALARIE=PSA_SALARIE'+
                             ' WHERE (PSA_ETABLISSEMENT="'+Etab+'" OR ETB_RATTACHEHANDI="'+Etab+'")'+ Habilitation + //PT11
                             ' ORDER BY PSA_ETABLISSEMENT,PGH_SALARIE';
     end;
end ;

procedure TOF_DECLARATIONHANDI.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_DECLARATIONHANDI.OnArgument (S : String ) ;
var AnneeE,Et:String;
   A,B,C,D1,D2,check1,check2,check3,check4,I,Vente,Achat,Dem,Branche,Entreprise,Etablissement:TCheckBox;
   QET,Q:TQuery;
   Etab: THValComboBox ;
   LIB:THLabel;
   DateEffectif,EffPCS1,EffPCS2,EffPCS3,EffPCS4,EffPCS5,EffPCS6,EffPCS7,EffPCS8,TOTPCS,VART1,DateEdition,DateAgr:THEdit;
   CodePCS1,CodePCS2,CodePCS3,CodePCS4,CodePCS5,CodePCS6,CodePCS7,CodePCS8:THEdit;
   Num : Integer;
begin
     Inherited ;
    //DEB PT11
    Habilitation := '';
    if Assigned(MonHabilitation) and (MonHabilitation.LeSQL<>'') then
      Habilitation := ' AND ' + MonHabilitation.LeSQL;
    //FIN PT11
     Q:=OpenSQL('SELECT * FROM EXERSOCIAL WHERE PEX_ACTIF="X" ORDER BY PEX_ANNEEREFER DESC',TRUE) ;
If Q.eof then                                               // PortageCWAS
   begin
   PGIBox('Vous n''avez pas d''exercice social actif',Ecran.Caption);
   Exit;
   end;
Q.Next ;
if Not Q.EOF then
   begin
   AnneeE := Q.FindField ('PEX_ANNEEREFER').AsString;
   end
Else      //PT- 6
   begin
   Q.First;
   AnneeE:=IntToStr((Q.FindField ('PEX_ANNEEREFER').AsInteger)-1);
   end;
   Ferme(Q);
   SetControlText('XX_VARANNEE',AnneeE);
     SetControlCaption('LIBI1','Si l''établissement a eu en permanence moins de 20 salariés en '+AnneeE+' remplir les champs suivants '); //PT- 5
     A:=TCheckBox(GetControl('XX_VAR3A'));
     B:=TCheckBox(GetControl('XX_VAR3B'));
     C:=TCheckBox(GetControl('XX_VAR3C'));
     D1:=TCheckBox(GetControl('XX_VAR3D'));
     D2:=TCheckBox(GetControl('XX_VAROBJ3D'));
     I:=TCheckBox(GetControl('LIBI2'));
     A.caption:='L''établissement a employé des bénéficiaires en ' +AnneeE;
     B.caption:='L''établissement a signé un des contrats suivants en ' +AnneeE; //PT- 5
     C.caption:='L''établissement a versé une contribution à l''AGEFIPH au titre de la présente déclaration ';
     D1.caption:='L''établissement est concerné par un accord en vigueur en '+AnneeE;
     D2.caption:='En cas d''accord, les objectifs fixés pour '+AnneeE+' ont été atteints';
     I.caption:='Combien de salariés au 31 décembre '+ AnneeE;
     QET:=OpenSQL('SELECT ET_ETABLISSEMENT FROM ETABLISS LEFT JOIN ETABCOMPL ON ET_ETABLISSEMENT=ETB_ETABLISSEMENT WHERE  ETB_RATTACHEHANDI<>""',True); //PT10
     Et:='';
     If Not QET.Eof Then Et:=QET.findfield('ET_ETABLISSEMENT').AsString; // // PortageCWAS
     Ferme(QET);
     Etab:=THValComboBox(GetControl('XX_VARETAB'));
     If etab<>NIL Then Etab.OnChange:=effectifs;
     Etab.value:=Et;
     Check1 := TCheckBox(GetControl('XX_VAR3A'));
     if check1<>nil then Check1.OnClick:=partie3a;
     Check2 := TCheckBox(GetControl('XX_VAR3B'));
     if check2<>nil then Check2.onclick:=partie3b;
     Check3 := TCheckBox(GetControl('XX_VAR3C'));
     if check3<>nil then Check3.onclick:=partie3c;
     Check4 := TCheckBox(GetControl('XX_VAR3D'));
     if check4<>nil then Check4.onclick:=partie3d;
     EffPCS1:=THEdit(GetControl('XX_VAREFFPCS1'));
     If EffPCS1<>NIL Then EffPCS1.OnChange:=ChangePCS;
     CodePCS1:=THEdit(GetControl('XX_VARCODEPCS1'));
     If CodePCS1<>NIL Then CodePCS1.OnExit:=VerifPCS;
     EffPCS2:=THEdit(GetControl('XX_VAREFFPCS2'));
     If EffPCS2<>NIL Then EffPCS2.OnChange:=ChangePCS;
     CodePCS2:=THEdit(GetControl('XX_VARCODEPCS2'));
     If CodePCS2<>NIL Then CodePCS2.OnExit:=VerifPCS;
     EffPCS3:=THEdit(GetControl('XX_VAREFFPCS3'));
     If EffPCS3<>NIL Then EffPCS3.OnChange:=ChangePCS;
     CodePCS3:=THEdit(GetControl('XX_VARCODEPCS3'));
     If CodePCS3<>NIL Then CodePCS3.OnExit:=VerifPCS;
     EffPCS4:=THEdit(GetControl('XX_VAREFFPCS4'));
     If EffPCS4<>NIL Then EffPCS4.OnChange:=ChangePCS;
     CodePCS4:=THEdit(GetControl('XX_VARCODEPCS4'));
     If CodePCS4<>NIL Then CodePCS4.OnExit:=VerifPCS;
     EffPCS5:=THEdit(GetControl('XX_VAREFFPCS5'));
     If EffPCS5<>NIL Then EffPCS5.OnChange:=ChangePCS;
     CodePCS5:=THEdit(GetControl('XX_VARCODEPCS5'));
     If CodePCS5<>NIL Then CodePCS5.OnExit:=VerifPCS;
     EffPCS6:=THEdit(GetControl('XX_VAREFFPCS6'));
     If EffPCS6<>NIL Then EffPCS6.OnChange:=ChangePCS;
     CodePCS6:=THEdit(GetControl('XX_VARCODEPCS6'));
     If CodePCS6<>NIL Then CodePCS6.OnExit:=VerifPCS;
     EffPCS7:=THEdit(GetControl('XX_VAREFFPCS7'));
     If EffPCS7<>NIL Then EffPCS7.OnChange:=ChangePCS;
     CodePCS7:=THEdit(GetControl('XX_VARCODEPCS7'));
     If CodePCS7<>NIL Then CodePCS7.OnExit:=VerifPCS;
     EffPCS8:=THEdit(GetControl('XX_VAREFFPCS8'));
     If EffPCS8<>NIL Then EffPCS8.OnChange:=ChangePCS;
     CodePCS8:=THEdit(GetControl('XX_VARCODEPCS8'));
     If CodePCS8<>NIL Then CodePCS8.OnExit:=VerifPCS;
     TOTPCS:=THEdit(GetControl('XX_VART2NA'));
     If TotPCS<>NIL Then TotPCS.OnChange:=ChangeT2;
     VART1:=THEdit(GetControl('XX_VART1'));
     If VART1<>NIL Then VART1.OnChange:=ChangeT1;
     Vente := TCheckBox(GetControl('XX_VENTE'));
     if Vente<>nil then Vente.OnClick:=ChangeAdresse;
     Achat := TCheckBox(GetControl('XX_ACHAT'));
     if Achat<>nil then Achat.OnClick:=ChangeAdresse;
     Dem:= TCheckBox(GetControl('XX_DEM'));
     if Dem<>nil then Dem.OnClick:=ChangeAdresse;
     Branche:= TCheckBox(GetControl('XX_VARDB3D'));
     if Branche<>nil then Branche.OnClick:=ChangeAccord;
     Entreprise:= TCheckBox(GetControl('XX_VARDEN3D'));
     if Entreprise<>nil then Entreprise.OnClick:=ChangeAccord;
     Etablissement:= TCheckBox(GetControl('XX_VARDET3D'));
     if Etablissement<>nil then Etablissement.OnClick:=ChangeAccord;
     DateEdition:=THEdit(GetControl('XX_VARDATEJ'));
     If DateEdition<>NIL Then DateEdition.OnExit:=verifdateedition;
     If DateEdition<>NIL Then DateEdition.OnElipsisClick := DateElipsisclick;
     DateAgr:=THEdit(GetControl('XX_VARDATEA3D'));
     If DateAgr<>NIL Then DateAgr.OnExit:=verifdateagrement;
     If DateAgr<>NIL Then DateAgr.OnElipsisClick := DateElipsisclick;
     DateEffectif:=THEdit(GetControl('XX_NBSAL3112'));  // PT 4
     If DateEffectif<>Nil Then DateEffectif.OnExit:=VerifEffectif;
     if (VH_Paie.PGPCS2003) then  //PT- 7
     begin
        For Num := 1 to 8 do
        begin
                SetControlProperty ('XX_VARCODEPCS'+IntToStr(Num), 'datatype', 'PGCODEPCSESE') ;
        end;
     end;
end ;



procedure TOF_DECLARATIONHANDI.ChangeAccord(Sender: Tobject);
begin
if sender=nil then exit;
if TCheckBox(sender).Name='XX_VARDB3D' then
   begin
   if GetControlText('XX_VARDB3D')='X' then
      begin
      SetControlEnabled('XX_VARDEN3D',False);
      SetControlChecked('XX_VARDEN3D',False);
      SetControlEnabled('XX_VARDET3D',False);
      SetControlChecked('XX_VARDET3D',False);
      end
      Else
      begin
      SetControlEnabled('XX_VARDET3D',True);
      SetControlEnabled('XX_VARDEN3D',True);
      end;
   end;
if TCheckBox(sender).Name='XX_VARDEN3D' then
   begin
   if GetControlText('XX_VARDEN3D')='X' then
      begin
      SetControlEnabled('XX_VARDB3D',False);
      SetControlChecked('XX_VARDB3D',False);
      SetControlEnabled('XX_VARDET3D',False);
      SetControlChecked('XX_VARDET3D',False);
      end
      Else
      begin
      SetControlEnabled('XX_VARDB3D',True);
      SetControlEnabled('XX_VARDET3D',True);
      end;
   end;
if TCheckBox(sender).Name='XX_VARDET3D' then
   begin
   if GetControlText('XX_VARDET3D')='X' then
      begin
      SetControlEnabled('XX_VARDB3D',False);
      SetControlChecked('XX_VARDB3D',False);
      SetControlEnabled('XX_VARDEN3D',False);
      SetControlChecked('XX_VARDEN3D',False);
      end
      Else
      begin
      SetControlEnabled('XX_VARDB3D',True);
      SetControlEnabled('XX_VARDEN3D',True);
      end;
   end;
end;

procedure TOF_DECLARATIONHANDI.ChangeAdresse(Sender: Tobject);
begin
if sender=nil then exit;
if TCheckBox(sender).Name='XX_VENTE' then
   begin
   if GetControlText('XX_VENTE')='X' then
      begin
      SetControlEnabled('XX_ACHAT',False);
      SetControlChecked('XX_ACHAT',False);
      SetControlEnabled('XX_DEM',False);
      SetControlChecked('XX_DEM',False);
      end
      Else
      begin
      SetControlEnabled('XX_ACHAT',True);
      SetControlEnabled('XX_DEM',True);
      end;
   end;
if TCheckBox(sender).Name='XX_ACHAT' then
   begin
   if GetControlText('XX_ACHAT')='X' then
      begin
      SetControlEnabled('XX_VENTE',False);
      SetControlChecked('XX_VENTE',False);
      SetControlEnabled('XX_DEM',False);
      SetControlChecked('XX_DEM',False);
      end
      Else
      begin
      SetControlEnabled('XX_VENTE',True);
      SetControlEnabled('XX_DEM',True);
      end;
   end;
if TCheckBox(sender).Name='XX_DEM' then
   begin
   if GetControlText('XX_DEM')='X' then
      begin
      SetControlEnabled('XX_ACHAT',False);
      SetControlChecked('XX_ACHAT',False);
      SetControlEnabled('XX_VENTE',False);
      SetControlChecked('XX_VENTE',False);
      end
      Else
      begin
      SetControlEnabled('XX_VENTE',True);
      SetControlEnabled('XX_ACHAT',True);
      end;
   end;
end;


procedure TOF_DECLARATIONHANDI.effectifs(Sender: Tobject);
var  Etab,DDTEFP,AnneeE,CodePCS,SommeBenef:String;
   Etablissement: THValComboBox ;
   QPCS,QPCS2,Q,QT1,QDDT,Qbeneficiaire:TQuery;
   Tob_PCS,TPCS : TOB;
   T1,T2:double;
   ie,ic,CodePCSINT,CalculAssiette,T1AI,T2AI,i:integer;
   findecl,debdecl,DatF:TDateTime;
   Jour, Mois, Annee :word;
   STRic,STRie,FirstPCS,T1AS,T2AS,Si:String;
   result:boolean;
   T1A,T2A:real;
   TobSalarie : Tob;
   TrentD,TrentN : Integer;
   UnitePE,Montant,SommePCS : Double;
begin
SommePCS := 0;
Q:=OpenSQL('SELECT * FROM EXERSOCIAL WHERE PEX_ACTIF="X" ORDER BY PEX_ANNEEREFER DESC',TRUE) ;
If Q.eof then exit; //// // PortageCWAS
Q.Next ;
if Not Q.EOF then
   begin
   DatF := Q.FindField ('PEX_FINPERIODE').AsFloat;//Q.Fields[7].AsFloat; // Recup date de fin periode en cours
   DecodeDate (DatF, Annee, Mois, Jour);
   AnneeE := Q.FindField ('PEX_ANNEEREFER').AsString;//Q.Fields[8].AsString; // recup Annee de exercice
   DebDecl := Q.FindField ('PEX_DATEDEBUT').AsDateTime;
   FinDecl := Q.FindField ('PEX_DATEFIN').AsDateTime;
   end
Else  //PT- 6
   begin
   Q.First;
   AnneeE := InttoStr((Q.FindField ('PEX_ANNEEREFER').AsInteger)-1);
   DebDecl := PlusDate(Q.FindField ('PEX_DATEDEBUT').AsDateTime,-1,'A');
   FinDecl := PlusDate(Q.FindField ('PEX_DATEFIN').AsDateTime,-1,'A');  
   end;
Ferme (Q);
SetControlText('XX_VARANNEE',AnneeE);

//FIN RECUP ANNE EXERCICE PRECEDENT

  Etablissement:=THValComboBox(GetControl('XX_VARETAB'));
  Etab:=Etablissement.value;
  DDTEFP:='';
  QDDT:=OpenSQL('SELECT ETB_CODEDDTEFPGU FROM ETABCOMPL WHERE ETB_ETABLISSEMENT="'+etab+'"',True);
  If Not QDDT.Eof Then DDTEFP:=QDDT.findfield('ETB_CODEDDTEFPGU').AsString; //DB2 // PortageCWAS
  SetControlText('XX_VARDDTEFP',DDTEFP);
  Ferme(QDDT);

  T1:=0;
T2:=0;
  //DEBUT PT- 9
  QT1:=OpenSQL('SELECT SUM (PSA_UNITEPRISEFF)SOMMET1,SUM(PPU_DENOMINTRENT) DENTRENTIEME,SUM(PPU_NUMERATTRENT) NUTRENTIEME '+
                     'FROM SALARIES LEFT JOIN PAIEENCOURS ON PSA_SALARIE=PPU_SALARIE '+
                     ' LEFT JOIN ETABCOMPL ON PSA_ETABLISSEMENT=ETB_ETABLISSEMENT WHERE PSA_PRISEFFECTIF="X"'+ //PT10
                     ' AND (PSA_ETABLISSEMENT="'+Etab+'" OR ETB_RATTACHEHANDI="'+Etab+'") AND ((PSA_DATESORTIE IS NULL) OR (PSA_DATESORTIE>"'+UsDateTime(DebDecl)+'") OR (PSA_DATESORTIE="'+UsDateTime(IDate1900)+'"))'+
                     ' AND ((PSA_DATEENTREE="'+UsDateTime(IDate1900)+'") OR (PSA_DATEENTREE IS NULL) OR (PSA_DATEENTREE<"'+UsDateTime(FinDecl)+'"))'+
                     ' AND PPU_DATEDEBUT>="'+UsDateTime(DebDecl)+'" AND PPU_DATEFIN<="'+UsDateTime(FinDecl)+'" ' + Habilitation + ' GROUP BY PSA_SALARIE',True); //PT11
  TobSalarie := Tob.Create('CalculEffectif',Nil,-1);
  TobSalarie.LoadDetailDB('CalculEffectif','','',QT1,False);
  T1 := 0;
  Ferme(QT1);
  For i := 0 to TobSalarie.Detail.Count - 1 do
  begin
       If TobSalarie.Detail[i].GetValue('DENTRENTIEME') <> null then TrentD := TobSalarie.Detail[i].GetValue('DENTRENTIEME')
       else Trentd := 0;
       if TobSalarie.Detail[i].GetValue('NUTRENTIEME') <> null then TrentN := TobSalarie.Detail[i].GetValue('NUTRENTIEME')
       else TrentN := 0;
       If TobSalarie.Detail[i].GetValue('SOMMET1') <> null then UnitePE := TobSalarie.Detail[i].GetValue('SOMMET1')
       else UnitePE := 0;
       If TrentD <> 0 then Montant := ((TrentN / TrentD) * UnitePE) / 12
       else Montant := 0;
       T1 := T1 + Montant;
  end;
  TobSalarie.Free;
  //FIN PT- 9
   T1A:=int(T1);
   T1AS:=FloatToStr(T1A);
   SetControlText('XX_VART1',T1AS);
    For i:=1 to 8 do      //Initialisation des THEdit
        Begin
        Si:=IntToStr(i);
        SetcontrolText('XX_VARCODEPCS'+Si,'');
        SetControlText('XX_VAREFFPCS'+Si,'');
        end;
  ie:=1;ic:=1;
  QPCS:=OpenSQL('SELECT DISTINCT PSA_CODEEMPLOI FROM SALARIES LEFT JOIN ETABCOMPL ON ETB_ETABLISSEMENT=PSA_ETABLISSEMENT'+  //PT10
  ' WHERE (PSA_ETABLISSEMENT="'+Etab+'" OR ETB_RATTACHEHANDI="'+Etab+'") AND PSA_CODEEMPLOI<>""' + Habilitation,True); //PT10 //PT11
  Tob_PCS:=Tob.create('Les CodePCS',nil,-1);
  Tob_PCS.LoadDetailDB('SALARIES','','',QPCS,False);
  Ferme(QPCS);
  FirstPCS:='';
  TPCS:=Tob_PCS.FindFirst([''],[''],False);
  while TPCS<>NIL do
  begin
  result:=False;
  CodePCS:=TPCS.GetValue('PSA_CODEEMPLOI');
  if FirstPCS=CodePCS then Break;
  if FirstPCS='' then FirstPCS:=CodePCS;
//  If IsNumeric(CodePCS) then           //PT- 7
  if NOT VH_Paie.PGPCS2003 then
  begin
        CodePCSINT:=StrToInt(CodePCS);
          Case CodePCSINT of
                    3862,3863,4802,5223,5315,5317,5441,5445,5510,6234,6241,6242,6243,6244,6245,6246,6331,6341,6343,6411,6412,6415,6511,6512,6513,6522,6531,6532,6741,6742,6911,6921 : Result:=True;
           End;
  end
  else
  begin                                 //PT- 8
        if CodePCS = '389b' then result := true
        else if CodePCS = '389c' then result := true
        else if CodePCS = '480b' then result := true
        else if CodePCS = '526e' then result := true
        else if CodePCS = '533b' then result := true
        else if CodePCS = '533c' then result := true
        else if CodePCS = '534a' then result := true
        else if CodePCS = '534b' then result := true
        else if CodePCS = '546a' then result := true
        else if CodePCS = '546e' then result := true
        else if CodePCS = '546d' then result := true
        else if CodePCS = '553b' then result := true
        else if CodePCS = '624d' then result := true
        else if CodePCS = '621a' then result := true
        else if CodePCS = '621b' then result := true
        else if CodePCS = '621c' then result := true
        else if CodePCS = '621d' then result := true
        else if CodePCS = '621e' then result := true
        else if CodePCS = '621f' then result := true
        else if CodePCS = '621g' then result := true
        else if CodePCS = '632c' then result := true
        else if CodePCS = '632a' then result := true
        else if CodePCS = '632e' then result := true
        else if CodePCS = '641a' then result := true
        else if CodePCS = '644a' then result := true
        else if CodePCS = '641b' then result := true
        else if CodePCS = '643a' then result := true
        else if CodePCS = '651a' then result := true
        else if CodePCS = '651b' then result := true
        else if CodePCS = '654c' then result := true
        else if CodePCS = '652b' then result := true
        else if CodePCS = '654b' then result := true
        else if CodePCS = '656b' then result := true
        else if CodePCS = '656c' then result := true
        else if CodePCS = '671a' then result := true
        else if CodePCS = '671c' then result := true
        else if CodePCS = '671d' then result := true
        else if CodePCS = '691a' then result := true
        else if CodePCS = '692a' then Result:=True;
  end;
  If (result=true) then
     Begin
     //DEBUT PT- 9
      QT1:=OpenSQL('SELECT SUM (PSA_UNITEPRISEFF)SOMMET1,SUM(PPU_DENOMINTRENT) DENTRENTIEME,SUM(PPU_NUMERATTRENT) NUTRENTIEME '+
                     'FROM SALARIES LEFT JOIN PAIEENCOURS ON PSA_SALARIE=PPU_SALARIE'+
                     ' LEFT JOIN ETABCOMPL ON ETB_ETABLISSEMENT=PSA_ETABLISSEMENT'+  //PT10
                     ' where PSA_CODEEMPLOI="'+CodePCS+'" AND PSA_PRISEFFECTIF="X"'+
                     ' AND (PSA_ETABLISSEMENT="'+Etab+'"  OR ETB_RATTACHEHANDI="'+Etab+'") AND ((PSA_DATESORTIE IS NULL) OR (PSA_DATESORTIE>"'+UsDateTime(DebDecl)+'") OR (PSA_DATESORTIE="'+UsDateTime(IDate1900)+'"))'+
                     ' AND ((PSA_DATEENTREE="'+UsDateTime(IDate1900)+'") OR (PSA_DATEENTREE IS NULL) OR (PSA_DATEENTREE<"'+UsDateTime(FinDecl)+'"))'+
                     ' AND PPU_DATEDEBUT>="'+UsDateTime(DebDecl)+'" AND PPU_DATEFIN<="'+UsDateTime(FinDecl)+'"' + Habilitation + ' GROUP BY PSA_SALARIE',True); //PT11
         TobSalarie := Tob.Create('CalculEffectif',Nil,-1);
         TobSalarie.LoadDetailDB('CalculEffectif','','',QT1,False);
         T1 := 0;
         Ferme(QT1);
         For i := 0 to TobSalarie.Detail.Count - 1 do
         begin
              If TobSalarie.Detail[i].GetValue('DENTRENTIEME') <> null then TrentD := TobSalarie.Detail[i].GetValue('DENTRENTIEME')
               else Trentd := 0;
               if TobSalarie.Detail[i].GetValue('NUTRENTIEME') <> null then TrentN := TobSalarie.Detail[i].GetValue('NUTRENTIEME')
               else TrentN := 0;
               If TobSalarie.Detail[i].GetValue('SOMMET1') <> null then UnitePE := TobSalarie.Detail[i].GetValue('SOMMET1')
               else UnitePE := 0;
               If TrentD <> 0 then Montant := ((TrentN / TrentD) * UnitePE) / 12
               else Montant := 0;
               SOMMEPCS := SOMMEPCS + Montant;
       end;
       TobSalarie.Free;
       // FIN PT- 9
      If SOMMEPCS<> 0 Then
         begin
         STRic:=IntToStr(ic);
         SetcontrolText('XX_VARCODEPCS'+STRic,CodePCS);
         STRie:=IntToStr(ie);
         SetControlText('XX_VAREFFPCS'+STRie,FloatToStr(SOMMEPCS));
         SetControlEnabled('XX_VAREFFPCS'+STRie,True);
         ie:=ie+1;ic:=ic+1;
         T2:=T2 + SOMMEPCS;
         end;
      end;
  TPCS:=Tob_PCS.FindNext([''],[''],False);
  end;
  Tob_PCS.Free;
  T2A:=int(T2);
  T2AS:=FloatToStr(T2A);
  T1AI:=StrToInt(T1AS);
  T2AI:=StrToInt(T2AS);
  SetControlText('XX_VART2',T2AS);
  SetControlText('XX_VART2NA',FloatToStr(T2));
  CalculAssiette:=T1AI-T2AI;
  SetControlText('XX_VART1T2',IntToStr(CalculAssiette));

  If ExisteSQL('SELECT PGH_SALARIE FROM HANDICAPE Left Join SALARIES on PSA_SALARIE=PGH_SALARIE '+
               'LEFT JOIN ETABCOMPL ON ETB_ETABLISSEMENT=PSA_ETABLISSEMENT'+ //PT10
               ' WHERE PSA_DATEENTREE>="'+UsDateTime(DebDecl)+'" AND PSA_DATEENTREE<="'+UsDateTime(FinDecl)+'" AND (PSA_ETABLISSEMENT="'+Etab+'" OR ETB_RATTACHEHANDI="'+Etab+'")' + Habilitation ) Then //PT11
               Begin
                    SetControlText('XX_VAR3A','X');
                    Qbeneficiaire:=OpenSQL('SELECT SUM (PSA_UNITEPRISEFF)SOMME FROM SALARIES LEFT JOIN HANDICAPE ON PSA_SALARIE=PGH_SALARIE'+
                    ' LEFT JOIN ETABCOMPL ON ETB_ETABLISSEMENT=PSA_ETABLISSEMENT'+ //PT10
                    ' WHERE (PGH_HANDICAP="X" OR PGH_COTOREP="X") AND PSA_PRISEFFECTIF="X" AND (PSA_ETABLISSEMENT="'+Etab+'" OR ETB_RATTACHEHANDI="'+Etab+'")'+
                    ' AND PSA_DATEENTREE>="'+UsDateTime(DebDecl)+'" AND PSA_DATEENTREE<="'+UsDateTime(FinDecl)+'"' + Habilitation,TRUE); //PT11
                    SOMMEbenef:='';
                    If Not Qbeneficiaire.eof then SOMMEbenef:=Qbeneficiaire.FindField('SOMME').AsString;  // // PortageCWAS
                    SetControlText('XX_VARNB3A',SOMMEbenef);
                    Ferme(Qbeneficiaire);
               end
  Else
               Begin
               SetControlText('XX_VARNB3A','');
               SetControlText('XX_VAR3A','-');
               end;

end;


procedure TOF_DECLARATIONHANDI.OnClose ;
begin
  Inherited ;

end ;

procedure TOF_DECLARATIONHANDI.partie3a(Sender : Tobject);
begin
     Inherited ;
     if sender=nil then exit;
     if GetControlText('XX_VAR3A')='X' then
        Begin
        SetControlEnabled('XX_VARNB3A',True);
        end
     Else
        Begin
        SetControlEnabled('XX_VARNB3A',False);
        end;
end ;

procedure TOF_DECLARATIONHANDI.partie3b(Sender : Tobject);
begin
     Inherited ;
     if sender=nil then exit;
     if GetControlText('XX_VAR3B')='X' then
        Begin
        SetControlEnabled('XX_VARNB3B',True);
        end
     Else
        Begin
        SetControlEnabled('XX_VARNB3B',False);
        end;
end ;

procedure TOF_DECLARATIONHANDI.partie3c(Sender : Tobject);
var T1:String;
begin
     Inherited ;
     if sender=nil then exit;
     if GetControlText('XX_VAR3C')='X' then
        Begin
        SetControlEnabled('XX_VARNB3C',True);
        SetControlEnabled('XX_VARMONT3C',True);
        SetControlEnabled('XX_VAREFF3C',True);
        //PT- 2
        T1:=GetControlText('XX_VART1');
        SetControlText('XX_VAREFF3C',T1);
        end
     Else
        Begin
        SetControlEnabled('XX_VARNB3C',False);
        SetControlEnabled('XX_VARMONT3C',False);
        SetControlEnabled('XX_VAREFF3C',False);
        end;
end ;

procedure TOF_DECLARATIONHANDI.partie3d(Sender : Tobject);
begin
     Inherited ;
     if sender=nil then exit;
     if GetControlText('XX_VAR3D')='X' then
        Begin
        SetControlEnabled('XX_VARDB3D',True);
        SetControlEnabled('XX_VARDEN3D',True);
        SetControlEnabled('XX_VARDET3D',True);
        SetControlEnabled('XX_VARDATEA3D',True);
        SetControlEnabled('XX_VARDEPART3D',True);
        SetControlEnabled('XX_VAROBJ3D',True);
        end
     Else
        Begin
        SetControlEnabled('XX_VARDB3D',False);
        SetControlEnabled('XX_VARDEN3D',False);
        SetControlEnabled('XX_VARDET3D',False);
        SetControlEnabled('XX_VARDATEA3D',False);
        SetControlEnabled('XX_VARDEPART3D',False);
        SetControlEnabled('XX_VAROBJ3D',False);
        end;
end ;

procedure TOF_DECLARATIONHANDI.ChangePCS(Sender: Tobject);
var  i:integer;
     T2,T:Double;
     si,VAREFF:String;
begin
T2:=0;
For i:=1 to 8 do
    begin
    si:=IntToStr(i);
    VAREFF:=(GetControlText('XX_VAREFFPCS'+si));
    If VarEff<>'' Then
       Begin
       T:=StrToFloat(VAREFF);
       T2:=T2+T;
       end;
    end;
SetControlText('XX_VART2NA',FloatToStr(T2));
end;

procedure TOF_DECLARATIONHANDI.ChangeT2(Sender: Tobject);
var T2NA:Double;
    T2A:real;
    T1,Assiette,T2AI:Integer;
    T2AS,VART2NA:String;
begin
  VART2NA:=GetControlText('XX_VART2NA');
  If VART2NA<>'' Then
  Begin
  T2NA:=StrToFloat(VART2NA);
  T2A:=int(T2NA);
  T2AS:=FloatToStr(T2A);
  SetControlText('XX_VART2',T2AS);
  T1:=StrToInt(GetControlText('XX_VART1'));
  T2AI:=StrToInt(T2AS);
  Assiette:=T1-T2AI;
  SetControlText('XX_VART1T2',IntToStr(Assiette));
  If ((T1<T2A) AND (T1<>0)) Then
     begin
     PGIBox('Attention, l''effectif T1 ne peut être infèrieur à l''effectif T2','Calcul de l''assiette d''assujetissement');
     SetControlEnabled('BValider',False);
     end
  Else SetControlEnabled('BValider',True);
  end;
end;

procedure TOF_DECLARATIONHANDI.ChangeT1(Sender: Tobject);
var T1,T2,Assiette:integer;
Begin
If GetControlText('XX_VART1')<>'' Then
   begin
   T1:=StrToInt(GetControlText('XX_VART1'));
   T2:=StrToInt(GetControlText('XX_VART2'));
   Assiette:=T1-T2;
   SetControlText('XX_VART1T2',IntToStr(Assiette));
   AccesCondition;
   If ((T1<T2) AND (T1<>0)) Then
     begin
     PGIBox('Attention, l''effectif T1 ne peut être infèrieur à l''effectif T2','Calcul de l''assiette d''assujetissement');
     SetControlEnabled('BValider',False);
     end
   Else SetControlEnabled('BValider',True);
   end;
End;

procedure TOF_DECLARATIONHANDI.AccesCondition;
var Effectif:Integer;
begin
If GetControlText('XX_VART1')='' Then Exit;
Effectif:=StrToInt(GetControlText('XX_VART1'));
If Effectif<20 Then
   begin
   SetControlEnabled('XX_NBSAL3112',True);
   SetControlEnabled('XX_FERMDEF',True);
   end
Else
    begin
    SetControlText ('XX_NBSAL3112','');
    SetControlEnabled ('XX_NBSAL3112',False);
    SetControlEnabled ('XX_FERMDEF',False);
    SetControlChecked('XX_FERMDEF',False);
    end;
end;

//PT- 1 
procedure TOF_DECLARATIONHANDI.verifdateedition(Sender: Tobject);
var DateEdition:THEdit;
begin
DateEdition:=THEdit(GetControl('XX_VARDATEJ'));
If Not IsValidDate(DateEdition.text) Then
   begin
   PGIBox(''''+DateEdition.text+''' n''est pas une date correcte','CEGID PAIE S5');
   DateEdition.SetFocus ;
   exit;
   end;
end;

procedure TOF_DECLARATIONHANDI.verifdateagrement(Sender: Tobject);
var DateAgr:THEdit;
begin
DateAgr:=THEdit(GetControl('XX_VARDATEA3D'));
If Not IsValidDate(DateAgr.text) Then
   begin
   PGIBox(''''+DateAgr.text+''' n''est pas une date correcte','CEGID PAIE S5');
   DateAgr.SetFocus ;
   exit;
   end;
end;

procedure TOF_DECLARATIONHANDI.DateElipsisclick(Sender: TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;
//PT- 3
procedure TOF_DECLARATIONHANDI.VerifPCS(Sender: Tobject);
var  CodePCS:String;
     CodePCSINT,i:Integer;
     result:boolean;
begin
Result:=False;
CodePCS:=THEdit(Sender).text;
If CodePCs<>'' Then
   begin
   If IsNumeric(CodePCS) then    //PT- 7
        begin
   CodePCSINT:=StrToInt(CodePCS);
          Case CodePCSINT of
          3862,3863,4802,5223,5315,5317,5441,5445,5510,6234,6241,6242,6243,6244,6245,6246,6331,6341,6343,6411,6412,6415,6511,6512,6513,6522,6531,6532,6741,6742,6911,6921 : Result:=True;
           End;
        end;
   If (result=fALSE) and (IsNumeric(CODEPCS)) then
      begin
      PGIBOX('Ce code PCS ne fait pas partie de la liste','Emplois exigeants des conditions d''aptitudes particulières');
      THEdit(Sender).SetFocus ;
      end
   Else
       begin
       For i:=1 to 8 do
           begin
           If THEdit(Sender).Name='XX_VARCODEPCS'+IntToStr(i) Then
              begin
              SetControlEnabled('XX_VAREFFPCS'+IntToStr(i),True);
              end;
           end;
       end;
   end
   Else
       begin
       For i:=1 to 8 do
           begin
           If THEdit(Sender).Name='XX_VARCODEPCS'+IntToStr(i) Then
              begin
              SetControlEnabled('XX_VAREFFPCS'+IntToStr(i),False);
              SetControlText('XX_VAREFFPCS'+IntToStr(i),'');
              end;
           end;
       end;
end;

procedure TOF_DECLARATIONHANDI.VerifEffectif(Sender:TObject);    // PT 4
begin
If IsNumeric(GetControlText('XX_NBSAL3112')) Then
   begin
   If StrToInt(GetControlText('XX_NBSAL3112'))>20 Then
      begin
      PgiBox('L''effectif au 31/12 ne peut être supérieur à 20','Déclaration des travailleurs handicapés');
      SetControlText('XX_NBSAL3112','');
      end;
   end;
end;

Initialization

  registerclasses ( [ TOF_DECLARATIONHANDI ] ) ; 
end.
