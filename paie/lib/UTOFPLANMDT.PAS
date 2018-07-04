{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 22/06/2001
Modifié le ... : 30/08/2001
Description .. : Source TOF edition du planning des visistes médicales
Mots clefs ... : PAIE;MEDICINETRAVAIL
*****************************************************************
PT- 1 : 10/10/2001 : JL 562: Procédures de vérification des dates + calendrier
PT- 2 : 10/09/2003 : JL V_42 FQ 10777 Ajout des AGLLancefiche pour affichage salarié de l'établissement sélectionné ds mul
PT- 3 : 29/06/2004 : JL V_50 FQ 11296 ajout gestion apte sous réserve
PT- 4 : 28/02/2005 : JL V_60 FQ 12031 Rafraichissement de la liste
PT- 5 : 28/04/2005 : JL V_60 FQ 12233 Correction nom de champ Apte sous réserve si liste des visites non effectuées
PT- 6   V_65 04/03/2006 JL : FQ 12893 Ajout champs libres pour gestion surveillance médicale renforcée
PT- 7 : 27/04/2007 : JL V_720 Gestion accès depuis ribbonbar salarié
PT- 8 : 31/07/2007 : JL V_80 FQ 14633 Création depuis fiche salarié
PT9 19/12/2007 V_81 FC Concept accessibilité fiche salarié
PT10 26/09/2009 JS Formatage du matricule FQ n°15096
}
Unit UTOFPLANMDT ;

Interface

Uses  Controls, Classes, sysutils,
{$IFNDEF EAGLCLIENT}
     db, HDB,Fe_Main,
{$ELSE}
     eMul,MainEAGL,
{$ENDIF}
     HCtrls, HEnt1, HMsgBox, UTOF ,HTB97,
     ParamDat, P5Def, HQry,
     EntPaie,
     PGOutils2;//PT10

Type
  TOF_PLANMDT = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    Action:String; //PT9
    procedure VerifDate1(Sender:TObject);
    procedure VerifDate2(Sender:TObject);
    procedure DateElipsisclick(Sender: TObject);
    procedure GrilleDblClick (Sender : TObject);
    procedure CreerVisite (Sender : TObject);
    procedure ExitEdit(Sender: TObject); //PT10
  end ;

Implementation

// On remplit par défaut les critères sur les composants XX_WHERE
procedure TOF_PLANMDT.OnLoad ;
var
   S:String;
begin
     Inherited ;
     if GetControlText('RESULTAT')='TOUS' then
     SetControlText('XX_WHERERES','');
     if GetControlText('RESULTAT')='VNE' then
        begin
        S:= '(PVM_APTE="" OR PVM_APTE="-" )'+
        'AND (PVM_APTESR="" OR PVM_APTESR="-" )'+
        'AND (PVM_INAPTE="" OR PVM_INAPTE="-" )'+
        'AND (PVM_RECLASSE="" OR PVM_RECLASSE="-")';
        SetControlText('XX_WHERERES',S);
        end;
     if GetControlText('RESULTAT')='A' then
        SetControlText('XX_WHERERES','PVM_APTE="X"');
     if GetControlText('RESULTAT')='I' then
        SetControlText('XX_WHERERES','PVM_INAPTE="X"');
     if GetControlText('RESULTAT')='R' then
        SetControlText('XX_WHERERES','PVM_RECLASSE="X"');
     if GetControlText('RESULTAT')='ASR' then  //PT- 3
        SetControlText('XX_WHERERES','PVM_APTESR="X"');
     if GetControlText('TYPE')='TOUS' then
        SetControlText('XX_WHERETYPE','');
     if GetControlText('TYPE')='E' then
        SetControlText('XX_WHERETYPE','PVM_TYPEVISITMED="2"');
     if GetControlText('TYPE')='R' then
        SetControlText('XX_WHERETYPE','PVM_TYPEVISITMED="1"');
     if GetControlText('TYPE')='A' then
        SetControlText('XX_WHERETYPE','PVM_TYPEVISITMED="3"');
end;

procedure TOF_PLANMDT.OnArgument (S : String ) ;
var Num : Integer;
    Date1,Date2: THEdit;
    {$IFNDEF EAGLCLIENT}
    Grille : THDBGrid;
    {$ELSE}
    Grille : THGrid;
    {$ENDIF}
    BOuv : TToolBarButton97;
    Salarie : String;
     Edit : THEdit;//PT10
begin
     Inherited ;
     Salarie := ReadTokenPipe(S,';');
     If Salarie <> '' then     //PT- 7
     begin
      SetControlEnabled('PVM_SALARIE',False);
      SetControlText('PVM_SALARIE',Salarie);
     end;
     //DEB PT9
     Action := '';
     if S <> '' then
       Action := ReadTokenPipe(S,';');
     //FIN PT9
     For Num := 1 to VH_Paie.PGNbreStatOrg do
         begin
         if Num >4 then Break;
         VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)),GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));
         end;
     VisibiliteStat (GetControl ('PSA_CODESTAT'),GetControl ('TPSA_CODESTAT')) ;
     // DEBUT PT- 6
        For Num := 1 to VH_Paie.PgNbCombo do
        begin
             if Num >4 then Break;
             VisibiliteChampLibreSal(IntToStr(Num),GetControl ('PSA_LIBREPCMB'+IntToStr(Num)),GetControl ('TPSA_LIBREPCMB'+IntToStr(Num)));
        end;
        // FIN PT- 6
     SetControlText('RESULTAT','TOUS');
     SetControlText('TYPE','TOUS');
     Date1:=THEdit(GetControl('PVM_DATEVISITE'));
     If Date1<>NIL Then Date1.OnExit:=VerifDate1;
     If Date1<>NIL Then Date1.OnElipsisClick := DateElipsisclick;
     Date2:=THEdit(GetControl('PVM_DATEVISITE_'));
     If Date2<>NIL Then Date2.OnExit:=VerifDate2;
     If Date2<>NIL Then Date2.OnElipsisClick := DateElipsisclick;
     {$IFNDEF EAGLCLIENT}
     Grille := THDBGrid (GetControl ('Fliste'));
     {$ELSE}
     Grille := THGrid (GetControl ('Fliste'));
     {$ENDIF}
     if Grille  <>  NIL then Grille.OnDblClick  :=  GrilleDblClick;
     BOuv  :=  TToolbarButton97 (getcontrol ('BInsert'));
     if Bouv  <>  NIL then  BOuv.OnClick  :=  CreerVisite;
  //debut PT10
   Edit := THEdit(GetControl('PVM_SALARIE'));
      If Edit <> Nil then Edit.OnExit := ExitEdit;
  //fin PT10
     //DEB PT9
     if Action='ACTION=CONSULTATION' then
       SetControlProperty('BInsert','enabled',false);
     //FIN PT9
end ;

// PT- 1
procedure TOF_PLANMDT.VerifDate1(Sender:TObject);
var THDate1,THDate2:THEdit;
    Date1,Date2:TDatetime;
begin
THDate1:=THEdit(GetControl('PVM_DATEVISITE'));
If Not IsValidDate(THDate1.text) Then
   begin
   PGIBox(''''+THDate1.text+''' n''est pas une date correcte','CEGID PAIE S5');
   THDate1.SetFocus ;
   exit;
   end
Else
    begin
    THDate2:=THEdit(GetControl('PVM_DATEVISITE_'));
    Date1:=StrToDateTime(THDate1.text);
    Date2:=StrToDatetime(THDate2.text);
    If Date1>Date2 Then
       begin
       PGIBox('La date saisie doit être inférieure au '''+THDate2.Text+'''','Edition médecine du travail');
       THDate1.SetFocus ;
       end;
    end;
end;

procedure TOF_PLANMDT.VerifDate2(Sender:TObject);
var THDate1,THDate2:THEdit;
    Date1,Date2:TDatetime;
begin
THDate2:=THEdit(GetControl('PVM_DATEVISITE_'));
If Not IsValidDate(THDate2.text) Then
   begin
   PGIBox(''''+THDate2.text+''' n''est pas une date correcte','CEGID PAIE S5');
   THDate2.SetFocus ;
   exit;
   end
Else
    begin
    THDate1:=THEdit(GetControl('PVM_DATEVISITE'));
    Date1:=StrToDateTime(THDate1.text);
    Date2:=StrToDatetime(THDate2.text);
    If Date1>Date2 Then
       begin
       PGIBox('La date saisie doit être supérieure au '''+THDate1.Text+'''','Edition médecine du travail');
       THDate2.SetFocus ;
       end;
    end;
end;


procedure TOF_PLANMDT.DateElipsisclick(Sender: TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;

//DEBUT PT- 2
procedure TOF_PLANMDT.GrilleDblClick (Sender : TObject);
var St : String;
    Q_Mul : THQuery ;
    {$IFDEF EAGLCLIENT}
    Liste : THGrid;
    {$ENDIF}
    Bt : TToolBarButton97;
begin
        {$IFDEF EAGLCLIENT}
        Liste := THGrid(GetControl('FLISTE'));
        TFmul(Ecran).Q.TQ.Seek(Liste.Row-1) ;
        {$ENDIF}
        Q_Mul := THQuery(Ecran.FindComponent('Q'));
        If Q_MUL.FindField('PVM_SALARIE').AsString = '' Then
        begin
                PGIBox('Vous devez choisir un salarié',Ecran.Caption);
                Exit;
        end;
        st  := Q_MUL.FindField('PVM_SALARIE').AsString+';'+DateToStr(Q_MUL.FindField('PVM_DATEVISITE').AsDatetime);
        if (not JaiLeDroitTag(200076)) and (Action='ACTION=CONSULTATION') then //PT9
          AglLanceFiche('PAY','VISITEMEDT','',St,'ACTION=CONSULTATION')
        else
          AglLanceFiche('PAY','VISITEMEDT','',St,'');
        Bt  :=  TToolbarButton97 (GetControl('BCherche'));  //PT- 4
        if Bt  <>  NIL then Bt.click;
end;

procedure TOF_PLANMDT.CreerVisite(Sender : TObject);
var St,Sal : String;
    Bt : TToolBarButton97;
begin
        st  := GetControlText('PSA_ETABLISSEMENT');
        Sal := GetControlText('PVM_SALARIE');//PT- 8
        AglLanceFiche ('PAY','VISITEMEDT', '', '' ,st+';'+Sal+';ACTION=CREATION');//PT- 8
        Bt  :=  TToolbarButton97 (GetControl('BCherche'));
        if Bt  <>  NIL then Bt.click;
end;
//FIN PT- 2
//debut PT10
procedure TOF_PLANMDT.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
  if ((VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text))) then
      edit.text:=AffectDefautCode(edit,10);
end;
//fin PT10
Initialization
  registerclasses ( [ TOF_PLANMDT ] ) ;
end.
