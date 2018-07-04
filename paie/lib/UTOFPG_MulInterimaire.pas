{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 06/02/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : MUL_INTERIMAIRE ()
Mots clefs ... : TOF;MUL_INTERIMAIRE
*****************************************************************
PT1 : 18/12/2002 : JL 591 Modification des libellés de la fiche, fiche de bug n° 10374
PT2 : 10/02/2004 : JL V_50 Gestion des salariés dans table interim pour RH compétences
PT3 : 24/08/2004 : JL V_50 FQ 11536 Maj liste, gestion filte sur type intérim
PT4 : 25/06/2008 : FC V_850 FQ 15423 CWAS : impossible d'accéder à un intérimaire en saisie sans indiquer son matricule
}
Unit UTOFPG_MulInterimaire ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Mul,Hqry,DBGrids,HDB,FE_Main,
{$ELSE}
     emul,UtileAGL,MaineAgl,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,EntPaie,P5Def,ParamDat,PgOutils,PGoutils2,LookUp,HTB97 ;

Type
  TOF_MUL_INTERIMAIRE = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    Grille : String;
    procedure DateElipsisclick(Sender: TObject);
    procedure PresenceClick(Sender:TObject);
    procedure ExitEdit(Sender: TObject);
    procedure ChangeType(Sender:TObject);
    procedure CreerEnregistrement (Sender : TObject);
    procedure AccesEnregistrement (Sender : TObject);
  end ;

Implementation

procedure TOF_MUL_INTERIMAIRE.OnLoad ;
var St1,StTrav,StAnd,Where:String;
    DateP:TDateTime;
    THdefaut:THValComboBox;
    i:Integer;
    CPresent:TCheckBox;
begin
  Inherited ;
If Grille='EMP' Then
   begin
   CPresent:=TCheckBox(GetControl('CPRESENCE'));
   St1:=''; StTrav:=''; StAnd:='';
   If (CPresent<>Nil) and (CPresent.Checked=True) Then
      begin
      If Not IsValidDate(GetControlText('EDATEPRESENCE')) Then
         begin
         PgiBox('La date saisie n''est pas valide','Multicritères intérimaires');
         Exit;
         end;
      DateP:=StrToDate(GetControlText('EDATEPRESENCE'));
      St1:=' PEI_DEBUTEMPLOI<="'+UsDateTime(DateP)+'" AND PEI_FINEMPLOI>="'+UsDateTime(DateP)+'"';
      end;
   For i:=1 to VH_Paie.PGNbreStatOrg do
       begin
       If i>4 then break;
       THdefaut:=THValComboBox(GeTControl('TRAVAILN'+IntToStr(i)));
       If (THDefaut<>Nil) and (THDefaut.Value<>'') Then
          begin
          If St1<>'' Then StAnd:=' AND ';
          If StTrav<>'' Then StTrav:=StTrav+' AND PEI_TRAVAILN'+IntToStr(i)+'="'+THDefaut.Value+'"'
          Else StTrav:=' PEI_TRAVAILN'+IntToStr(i)+'="'+THDefaut.Value+'"';
          end;
       end;
   THdefaut:=THValComboBox(GeTControl('CODESTAT'));
   If (THDefaut<>Nil) and (THDefaut.Value<>'') Then
      begin
      If St1<>'' Then StAnd:=' AND';
      If StTrav<>'' Then StTrav:=StTrav+' AND PEI_CODESTAT="'+THDefaut.Value+'"'
      Else StTrav:=' PEI_CODESTAT="'+THDefaut.Value+'"';
      end;
  If (St1<>'') OR (StTrav<>'') Then Where:='PSI_TYPEINTERIM<>"CAN" AND PSI_TYPEINTERIM<>"SAL" AND PSI_INTERIMAIRE=(SELECT PEI_INTERIMAIRE FROM EMPLOIINTERIM'+
   ' WHERE PSI_INTERIMAIRE=PEI_INTERIMAIRE AND '+St1+StAnd+StTrav+')'
//   If (St1<>'') OR (StTrav<>'') Then Where:=St1+StAnd+StTrav
   Else Where:='PSI_TYPEINTERIM<>"CAN" AND PSI_TYPEINTERIM<>"SAL"';
   SetControlText('XX_WHERE',Where);
   end;
If Grille='RECUP' Then SetControlText('XX_WHERE','PSI_TYPEINTERIM<>"CAN" AND PSI_TYPEINTERIM<>"SAL" AND PSI_INTERIMAIRE NOT IN (SELECT PSA_SALARIE FROM SALARIES)');
end ;

procedure TOF_MUL_INTERIMAIRE.OnArgument (S : String ) ;
Var Num:Integer;
    CPresent:TCheckBox;
    EDateP,Edit:THEdit;
    TypeInterim:THValComboBox;
    {$IFNDEF EAGLCLIENT}
    Liste : THDBGrid;
    {$ELSE}
    Liste : THGrid;
    {$ENDIF}
begin
  Inherited ;
CPresent:=TCheckBox(GetControl('CPRESENCE'));
If CPresent<>Nil Then CPresent.OnClick:=PresenceClick;
EDateP:=THEdit(GetControl('EDATEPRESENCE'));
SetControlEnabled('EDATEPRESENCE',False);
If EDateP<>NIL Then EDateP.OnElipsisClick := DateElipsisclick;
Grille:=(ReadTokenSt(S));
If Grille='INT' Then        //PT1
   begin
   TFMul(Ecran).Caption:='Saisie';
   UpdateCaption(TFMul(Ecran));
   SetControlProperty('PSI_TYPEINTERIM','Plus','AND (CO_CODE="STA" OR CO_CODE="INT")');  //PT3
   SetControlText('XX_WHERE','PSI_TYPEINTERIM<>"CAN" AND PSI_TYPEINTERIM<>"SAL"');
   end;
If Grille='EMP' Then
   begin
   SetControlVisible('BInsert',False);
   SetControlVisible('CPRESENCE',True);
   SetControlVisible('EDATEPRESENCE',True);
   SetControlProperty('PSI_TYPEINTERIM','Plus','AND (CO_CODE="STA" OR CO_CODE="INT")');  //PT3
   For Num := 1 to VH_Paie.PGNbreStatOrg do
       begin
       if Num >4 then Break;
       VisibiliteChampSalarie (IntToStr(Num),GetControl ('TRAVAILN'+IntToStr(Num)),GetControl ('T_TRAVAILN'+IntToStr(Num)));
       end;
   VisibiliteStat (GetControl ('CODESTAT'),GetControl ('T_CODESTAT')) ;
   TFMul(Ecran).Caption:='Emplois des intérimaires et stagiaires';       //PT1
   UpdateCaption(TFMul(Ecran));
   end;
If (Grille='INT') OR (Grille='RECUP') Then SetControlProperty('PComplement','TabVisible',False);
If Grille='RECUP' Then
   begin
   SetControlVisible('BINSERT',False);
   TFMul(Ecran).Caption:='Choix de l''intérimaire à récupérer';
   UpdateCaption(TFMul(Ecran));
   end;
Edit:=THEdit(GetControl('PSI_INTERIMAIRE'));
If Edit<>Nil Then Edit.OnExit:=ExitEdit;
TypeInterim:=THValComboBox(GetControl('PSI_TYPEINTERIM'));
If TypeInterim<>Nil Then TypeInterim.OnChange:=ChangeType;
If Grille ='CANDIDATS' then
begin
        SetControlVisible('CPRESENCE',False);
        SetControlVisible('EDATEPRESENCE',False);
        SetControlEnabled('PSI_TYPEINTERIM',False);
        SetControlProperty('PSI_TYPEINTERIM','Plus','');
        SetControlText('PSI_TYPEINTERIM','CAN');
        TFMul(Ecran).Caption:='Liste des candidats';
        UpdateCaption(TFMul(Ecran));
        SetControlText('XX_WHERE','');
end;
If Grille ='EXTERIEUR' then
begin
        SetControlVisible('CPRESENCE',False);
        SetControlVisible('EDATEPRESENCE',False);
        SetControlEnabled('PSI_TYPEINTERIM',False);
        SetControlProperty('PSI_TYPEINTERIM','Plus','');
        SetControlText('PSI_TYPEINTERIM','EXT');
        TFMul(Ecran).Caption:='Liste des intervenants';
        UpdateCaption(TFMul(Ecran));
        SetControlText('XX_WHERE','');
end;
TFMul(Ecran).BInsert.OnClick := CreerEnregistrement;
{$IFNDEF EAGLCLIENT}
Liste := THDBGrid (GetControl ('Fliste'));
{$ELSE}
Liste := THGrid (GetControl ('Fliste'));
{$ENDIF}
If Liste <> Nil then Liste.OnDblClick := AccesEnregistrement;
end ;

procedure TOF_MUL_INTERIMAIRE.DateElipsisclick(Sender: TObject);
var key : char;
begin
key := '*';
ParamDate (Ecran, Sender, Key);
end;

procedure TOF_MUL_INTERIMAIRE.PresenceClick(Sender:TObject);
begin
If TCheckBox(Sender).Checked=True Then
   begin
   SetControlEnabled('EDATEPRESENCE',True);
   SetControlText('EDATEPRESENCE',DateToStr(Date));
   end
Else SetControlEnabled('EDATEPRESENCE',False);
end;

procedure TOF_MUL_INTERIMAIRE.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;

procedure TOF_MUL_INTERIMAIRE.ChangeType(Sender:TObject);
var Edit:THEdit;
begin
If Sender=Nil Then Exit;
Edit:=THEdit(GetControl('PSI_INTERIMAIRE'));
If Edit=Nil Then Exit;
If THValComboBox(Sender).Value='CAN' Then Edit.Plus:='CAN';
If THValComboBox(Sender).Value='STA' Then Edit.Plus:='STA';
If THValComboBox(Sender).Value='INT' Then Edit.Plus:='INT';
If THValComboBox(Sender).Value='SAL' Then Edit.Plus:='SAL';
If THValComboBox(Sender).Value='' Then Edit.Plus:='';
end;

procedure TOF_MUL_INTERIMAIRE.CreerEnregistrement (Sender : TObject);
var Bt : TToolBarButton97;
begin
        If Grille = 'CANDIDATS' then AglLanceFiche ('PAY','INTERIMAIRES', '', '' ,'CANDIDATS;ACTION=CREATION')
        else If Grille = 'EXTERIEUR' then AglLanceFiche ('PAY','INTERIMAIRES', '', '' ,'EXTERIEUR;ACTION=CREATION')
        else AglLanceFiche ('PAY','INTERIMAIRES', '', '' ,'INTERIMAIRES;ACTION=CREATION');
        Bt  :=  TToolbarButton97 (GetControl('BCherche'));
        if Bt  <>  NIL then Bt.click;
end;

procedure TOF_MUL_INTERIMAIRE.AccesEnregistrement (Sender : TObject);
var Bt : TToolBarButton97;
    Interimaire : String;
begin
{$IFDEF EAGLCLIENT}
//TFmul(Ecran).Q.TQ.Seek(Grille.Row-1) ;
  TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1); //PT4 Nécessaire pour se positionner sur la bonne ligne en cwas
{$ENDIF}
Interimaire := TFMul(Ecran).Q.FindField('PSI_INTERIMAIRE').AsString;
//DEBUT PT3
If Interimaire = '' then
begin
        PGIBox('Vous devez choisir un intérimaire',Ecran.Caption);
        Exit;
end;
//FIN PT3
if Grille = 'CANDIDATS' then AglLanceFiche ('PAY', 'INTERIMAIRES','',Interimaire,'CANDIDATS;ACTION=MODIFICATION');
if Grille = 'INT' then AglLanceFiche ('PAY', 'INTERIMAIRES','',Interimaire,'INTERIMAIRES;ACTION=MODIFICATION');
if Grille = 'EMP' then AglLanceFiche ('PAY', 'EMPLOIINTERIM_MUL','','', Interimaire+';MUL');
if Grille = 'RECUP' then  AglLanceFiche ('PAY', 'SALARIE', '','', 'ACTION=CREATION;INT;'+Interimaire);
if Grille = 'EXTERIEUR' then AglLanceFiche ('PAY', 'INTERIMAIRES','',Interimaire,'EXTERIEUR;ACTION=MODIFICATION');
Bt  :=  TToolbarButton97 (GetControl('BCherche'));  //PT3
if Bt  <>  NIL then Bt.click;
end;

Initialization
  registerclasses ( [ TOF_MUL_INTERIMAIRE ] ) ;
end.
