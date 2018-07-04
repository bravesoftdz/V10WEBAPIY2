{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 06/02/2002
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : EMPLOIINTERIM (EMPLOIINTERIM)
Mots clefs ... : TOM;EMPLOIINTERIM
*****************************************************************
}
{
PT1 17/12/2002 PH V591 Toutes les dates sont initialisées à idate1900 ou 2099 au lieu de null
PT2 18/12/2002 JL V591 Création dans annuaire
PT3 16/09/2003 JL V_42 Gestion nouveau codePCS             *
PT4 10/02/2004 JL V_50 Maj nouveau champs intérimaires gérés pour RH Compétences
PT5 25/08/2004 JL V_50 FQ 11547 Correction "erreur type variant incorrect" en validation
---- JL 20/03/2006 modification clé annuaire ----
}
Unit UtomEmploiInterimaires;

Interface

Uses Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fiche,Fe_Main,
{$ELSE}
     eFiche,eFichList,MainEAgl,
{$ENDIF}
     sysutils,HCtrls,HEnt1,HMsgBox,UTOM,UTob,EntPaie,P5Def,HTB97 ;

Type
  TOM_EMPLOIINTERIM = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;      //PT4
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    Private
    Interimaire,Action,TypeInt:String;
    procedure BtnAnnuIntClick(Sender:TObject);
    procedure BtnAnnuStaClick(Sender:TObject);
    end ;

Implementation

procedure TOM_EMPLOIINTERIM.OnNewRecord ;
Var NumOrdre:Integer;
    Q:TQuery;
begin
  Inherited ;
Q:=OpenSQL('SELECT MAX(PEI_ORDRE) AS MaxOrdre FROM EMPLOIINTERIM WHERE PEI_INTERIMAIRE="'+Interimaire+'"',True);
NumOrdre:=1;
if not Q.eof Then NumOrdre:=(Q.FindField('MaxOrdre').AsInteger)+1; // PortageCWAS
Ferme(Q);
SetField('PEI_ORDRE',NumOrdre);
SetField('PEI_INTERIMAIRE',Interimaire);
Q:=OpenSQL('SELECT PSI_CONFIDENTIEL FROM INTERIMAIRES WHERE PSI_INTERIMAIRE="'+Interimaire+'"',True);
If not Q.eof then SetField('PEI_CONFIDENTIEL',Q.FindField('PSI_CONFIDENTIEL').AsString);  // PortageCWAS
Ferme(Q);
SetField('PEI_AGENCEINTERIM',-1);
SetField('PEI_CENTREFORM',-1);
//PT1 17/12/2002 PH V591 Toutes les dates sont initialisées à idate1900 ou 2099 au lieu de null
SetField('PEI_DEBUTEMPLOI',IDate1900);
SetField('PEI_FINEMPLOI',IDate1900);
SetField('PEI_ESSAIDEBUT',IDate1900);
SetField('PEI_ESSAIFIN',IDate1900);
SetField('PEI_DATECONTRAT',IDate1900);
end ;

procedure TOM_EMPLOIINTERIM.OnAfterUpdateRecord ;     //PT4
var TobInterimaire,T : Tob;
    Q : TQuery;
    DateDeb,DateFin : TDateTime;
begin
  Inherited ;
If Not ExisteSQL('SELECT PEI_INTERIMAIRE FROM EMPLOIINTERIM WHERE PEI_DEBUTEMPLOI>"'+UsDateTime(GetField('PEI_DEBUTEMPLOI'))+'" AND PEI_INTERIMAIRE="'+GetField('PEI_INTERIMAIRE')+'"') then  //PT5
begin
        Q := OpenSQL('SELECT * FROM INTERIMAIRES WHERE PSI_INTERIMAIRE="'+GetField('PEI_INTERIMAIRE')+'"',True);
        TobInterimaire := Tob.Create('INTERIMAIRES',Nil,-1);
        TobInterimaire.LoadDetailDB('INTERIMAIRES','','',Q,False);
        Ferme(Q);
        T := TobInterimaire.FindFirst(['PSI_INTERIMAIRE'],[GetField('PEI_INTERIMAIRE')],False);
        If T <> Nil then
        begin
                DateDeb := GetField('PEI_DEBUTEMPLOI');
                DateFin := GetField('PEI_FINEMPLOI');
                T.PutValue('PSI_DATEENTREE',DateDeb);
                T.PutValue('PSI_DATESORTIE',DateFin);
                T.PutValue('PSI_ETABLISSEMENT',GetField('PEI_ETABLISSEMENT'));
                T.PutValue('PSI_LIBELLEEMPLOI',GetField('PEI_LIBELLEEMPLOI'));
                T.PutValue('PSI_QUALIFICATION',GetField('PEI_QUALIFICATION'));
                T.PutValue('PSI_COEFFICIENT',GetField('PEI_COEFFICIENT'));
                T.PutValue('PSI_CODESTAT',GetField('PEI_CODESTAT'));
                T.PutValue('PSI_TRAVAILN1',GetField('PEI_TRAVAILN1'));
                T.PutValue('PSI_TRAVAILN2',GetField('PEI_TRAVAILN2'));
                T.PutValue('PSI_TRAVAILN3',GetField('PEI_TRAVAILN3'));
                T.PutValue('PSI_TRAVAILN4',GetField('PEI_TRAVAILN4'));
                T.PutValue('PSI_LIBREPCMB1',GetField('PEI_LIBREPCMB1'));
                T.PutValue('PSI_LIBREPCMB2',GetField('PEI_LIBREPCMB2'));
                T.PutValue('PSI_LIBREPCMB3',GetField('PEI_LIBREPCMB3'));
                T.PutValue('PSI_LIBREPCMB4',GetField('PEI_LIBREPCMB4'));
                T.PutValue('PSI_INDICE',GetField('PEI_INDICE'));
                T.PutValue('PSI_NIVEAU',GetField('PEI_NIVEAU'));
                T.UpdateDB(False);
        end;
        TobInterimaire.Free;
end;
end;

procedure TOM_EMPLOIINTERIM.OnUpdateRecord ;
Var Mes:String;
    Debut,Fin,DebutE,FinE:Variant;
begin
  Inherited ;
Mes:='';
Debut:=GetField('PEI_DEBUTEMPLOI');
Fin:=GetField('PEI_FINEMPLOI');
DebutE:=GetField('PEI_ESSAIDEBUT');
FinE:=GetField('PEI_ESSAIFIN');
If GetField('PEI_ETABLISSEMENT')='' Then
   begin
   Mes:='#13#10- L''établissement';
   SetFocusControl('PEI_ETABLISSEMENT');
   end;
If Debut<10 Then
   begin
   Mes:=Mes+'#13#10- La date de début';
   SetFocusControl('PEI_DEBUTEMPLOI');
   end;
If Mes<>'' Then
   begin
   LastError:=1;
   PgiBox('La ou les informations suivantes sont obligatoires :'+Mes,'Saisie emploi');
   Exit;
   end;
If (Debut>Fin) and (Fin>10) Then
   begin
   Mes:='#13#10- La date de début ne peut être supérieur à la date de fin';
   SetFocusControl('PEI_DEBUTEMPLOI');
   end;
If DebutE>FinE Then
   begin
   Mes:=Mes+'#13#10- La date de début d''essai ne peut être supérieur à la date de fin d''essai';
   SetFocusControl('PEI_ESSAIDEBUT');
   end;
If (DebutE<Debut) AND (DebutE>10) Then
   begin
   Mes:=Mes+'#13#10- La date de début d''essai ne peut être inférieur à la date de début';
   SetFocusControl('PEI_ESSAIDEBUT');
   end;
If FinE>Fin Then
   begin
   Mes:=Mes+'#13#10- La date de fin d''essai ne peut être supérieur à la date de fin';
   SetFocusControl('PEI_ESSAIDEBUT');
   end;
If Mes<>'' Then
   begin
   LastError:=1;
   PgiBox('La ou les informations suivantes sont erronnées :'+Mes,'Saisie emploi');
   end;
end ;

procedure TOM_EMPLOIINTERIM.OnLoadRecord ;
var Num : Integer;
begin
  Inherited ;
If GetField('PEI_CODEEMPLOI')='' Then SetControlCaption('TPEI_CODEEMPLOI_','');
TFFiche(Ecran).Caption:='Saisie emploi '+RechDom('PGTYPEINTERIM',TypeInt,False)+' : '+
                   Interimaire+' '+RechDom('PGINTERIMAIRES',Interimaire,False);
TFFiche(Ecran).DisabledMajCaption := True;
For Num := 1 to VH_Paie.PGNbreStatOrg do
    begin
    if Num >4 then Break;
    VisibiliteChampSalarie (IntToStr(Num),GetControl ('PEI_TRAVAILN'+IntToStr(Num)),GetControl ('TPEI_TRAVAILN'+IntToStr(Num)));
    end;
VisibiliteStat (GetControl ('PEI_CODESTAT'),GetControl ('TPEI_CODESTAT')) ;
end ;

procedure TOM_EMPLOIINTERIM.OnChangeField ( F: TField ) ;
Var LibSal:THLabel;
begin
  Inherited ;
if (F.FieldName='PEI_LIBELLEEMPLOI') then
   begin
   If GetField('PEI_LIBELLEEMPLOI')='' Then SetControlCaption('TPEI_LIBELLEEMPLOI_','')
   Else SetControlCaption('TPEI_LIBELLEEMPLOI_',RechDom('PGLIBEMPLOI',GetField('PEI_LIBELLEEMPLOI'),False));
   end;
if (F.FieldName='PEI_QUALIFICATION') then
   begin
   If GetField('PEI_QUALIFICATION')='' Then SetControlCaption('TPEI_QUALIFICATION_','')
   Else SetControlCaption('TPEI_QUALIFICATION_',RechDom('PGLIBQUALIFICATION',GetField('PEI_QUALIFICATION'),False));
   end;
if (F.FieldName='PEI_COEFFICIENT') then
   begin
   If GetField('PEI_COEFFICIENT')='' Then SetControlCaption('TPEI_COEFFICIENT_','')
   Else SetControlCaption('TPEI_COEFFICIENT_',RechDom('PGLIBCOEFFICIENT',GetField('PEI_COEFFICIENT'),False));
   end;
If (F.FieldName='PEI_SALREMPL') Then
   begin
   LibSal:=THLabel(GetControl('LIBSALARIE'));
   If LibSal<>Nil Then
      begin
      If GetField('PEI_SALREMPL')='' Then LibSal.Caption:=''
      Else LibSal.Caption:=RechDom('PGSALARIE', GetField ('PEI_SALREMPL'),FALSE);
      end;
   end;
end ;

procedure TOM_EMPLOIINTERIM.OnArgument ( S: String ) ;
var  
     Q:TQuery;
         Num:Integer;
    {$IFNDEF EAGLCLIENT}
    Combo:THDbValComboBox;
    {$ELSE}
    Combo:THValComboBox;
    {$ENDIF}
    TLieu:THLabel;
    Numero:String;
    BtnAnnuInt,BtnAnnuSta:TToolBarButton97;
begin
  Inherited ;
Interimaire :=(ReadTokenSt(S));
Action:=trim(ReadTokenPipe(S,';'));
Q:=OpenSQL('SELECT PSI_TYPEINTERIM FROM INTERIMAIRES WHERE PSI_INTERIMAIRE="'+Interimaire+'"',True);
TypeInt:='';
if not Q.eof then TypeInt:=Q.FindField('PSI_TYPEINTERIM').AsString;   // PortageCWAS
Ferme(Q);
If TypeInt='STA' Then
   begin
   SetControlVisible('FINTERIMAIRES',False);
   SetControlCaption('TPEI_MONTANTCONTRAT','Montant des indemnités')
   end;
If TypeInt='INT' Then
   begin
   SetControlVisible('FSTAGIAIRES',false);
   SetControlCaption('TPEI_MONTANTCONTRAT','Montant global du contrat');
   end;
For num := 1 to VH_Paie.PgNbCombo do
  begin
  Numero:=InttoStr(num);
  if Num > VH_Paie.PgNbCombo then break;
{$IFNDEF EAGLCLIENT}
  Combo:=THDbValComboBox(GetControl('PEI_LIBREPCMB'+Numero));
{$ELSE}
  Combo:=THValComboBox(GetControl('PEI_LIBREPCMB'+Numero));
{$ENDIF}
  if Combo <> NIL then Combo.Visible :=TRUE;
  TLieu:=THLabel(GetControl('TPEI_LIBREPCMB'+Numero));
  if (TLieu <> NIL)   then
   begin
   TLieu.Visible :=TRUE;
   if Num = 1 then Begin TLieu.Caption :=VH_Paie.PgLibCombo1; end;
   if Num = 2 then Begin TLieu.Caption :=VH_Paie.PgLibCombo2; end;
   if Num = 3 then Begin TLieu.Caption :=VH_Paie.PgLibCombo3; end;
   if Num = 4 then Begin TLieu.Caption :=VH_Paie.PgLibCombo4; end;
   end;
  end;
SetControlProperty('PEI_MOTIFCTINT','DataType','PGMOTIFINTERIM');
BtnAnnuInt:=TToolBarButton97(Getcontrol('BTNANNUINT'));               // DEBUT PT2
If BtnAnnuInt<>Nil then BtnAnnuInt.OnClick:=BtnAnnuIntClick;
BtnAnnuSta:=TToolBarButton97(Getcontrol('BTNANNUSTA'));
If BtnAnnuSta<>Nil then BtnAnnuSta.OnClick:=BtnAnnuStaClick;         //FIN PT2
if VH_Paie.PGPCS2003 then SetControlproperty('PEI_CODEEMPLOI','DataType','PGCODEPCSESE');  // PT3
end ;

procedure TOM_EMPLOIINTERIM.BtnAnnuIntClick(Sender:TObject);                  //DEBUT PT2
begin
AGLLancefiche('YY','ANNUAIRE','','','ACTION=CREATION');
AvertirTable('PGAGENCEINTERIM');
end;

procedure TOM_EMPLOIINTERIM.BtnAnnuStaClick(Sender:TObject);
begin
AGLLancefiche('YY','ANNUAIRE','','','ACTION=CREATION');
AvertirTable('PGCENTREFORMATION');
end;
                                                                             //FIN PT2
Initialization
  registerclasses ( [ TOM_EMPLOIINTERIM ] ) ;
end.


