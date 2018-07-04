{***********UNITE*************************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 24/04/2003
Modifié le ... :   /  /    
Description .. : Gestion du multicritère et fiche de saisie des ressources de 
Suite ........ : type SAL
Mots clefs ... : PAIE;RESSOURECE
*****************************************************************}
{
PT1 14/02/2005 SB V_60 FQ 11349 Refonte effacement ressource
}
unit UTOFPGMulSalarie_res;

interface

uses  Classes,Controls,SysUtils,Utof,HCtrls,HQry,Hent1,HmsgBox,utob,HTB97,
      Ed_Tools,Vierge
{$IFNDEF EAGLCLIENT}
      ,HDB,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,mul
{$ELSE}
      ,MaineAGL,eMul
{$ENDIF}
     ,AGLInit;

Type
     { Multicritères}
     TOF_PGSalarie_res  = Class (TOF)
       procedure OnArgument (stArgument: String); Override;
       procedure OnLoad ;                         Override;
     private
       Q_Mul:THQuery ;
       BtnCherche : TToolbarButton97;
       procedure ExitEdit(Sender: TObject);
       Procedure GrilleDblClick(Sender : TObject);
       procedure OnClickAffectAuto (Sender : TObject);
     END ;
     { Fiche de saisie }
     TOF_PGRessource  = Class (TOF)
       procedure OnArgument (stArgument: String); Override;
       procedure OnLoad ;                         Override;
       procedure OnUpdate;                        override;
       procedure OnDelete;                        override;
       procedure OnClose;                         override;
     private
       Mode,Liste : String;
       Modifier,Supprimer : Boolean;
       QMul : TQUERY; // Query du mul
       {$IFNDEF EAGLCLIENT}
       procedure ClickNavigation (Sender: TObject);
       {$ENDIF}
       procedure ExitEdit        (Sender: TObject);
       procedure ChangeSal       (Sender: TObject);
       procedure ChangeRess      (Sender: TObject);
       procedure MajCaption;
       procedure EffaceZoneSalarie;
     END ;

const
	// libellés des messages
	TexteMessage: array[1..6] of string 	= (
          {1} 'Vous devez saisir un matricule existant.',
          {2} 'Vous devez saisir la ressource à affecter au salarié',
          {3} 'Vous devez saisir une ressource inexistante.',
          {4} 'Vous devez saisir un matricule non affecté. ',
          {5} 'Vous avez paramétré la reprise du matricule salarié dans la gestion des ressources.#13#10'+
              'Le code ressource et le matricule du salarié ne peut être différent.',
          {6} '');



implementation
uses EntPaie,PgOutils2,PgCommun;  

{ TOF_PGSalarie_res }


procedure TOF_PGSalarie_res.OnArgument(stArgument: String);
var
  {$IFDEF EAGLCLIENT}
  Grille : THGrid ;
  {$ELSE}
  Grille : THDBGrid;
  {$ENDIF}
  Btn : TToolbarButton97;
begin
  inherited;
  Q_Mul:=THQuery(Ecran.FindComponent('Q'));
  {$IFDEF EAGLCLIENT}
  Grille:=THGrid (GetControl ('Fliste'));
  {$ELSE}
  Grille:=THDBGrid (GetControl ('Fliste'));
  {$ENDIF}
  if Grille <> NIL then Grille.OnDblClick := GrilleDblClick;
  THEDit(GetControl('PSA_SALARIE')).OnExit:=ExitEdit;
  BtnCherche:=TToolbarButton97 (GetControl ('BCherche'));

  Btn:=TToolbarButton97 (GetControl ('TVALID'));
  if Btn<>nil then Btn.OnClick:=OnClickAffectAuto;
  
end;

procedure TOF_PGSalarie_res.OnLoad;
begin
  inherited;
  SetControlText('XX_WHERE','');
  SetControlVisible('TVALID',((GetControlText('CKSALNONAFFECT')='X') And (VH_PAie.PgTypeAffectRes<>'')));
  if GetControlText('CKSALNONAFFECT')='X' then
    Begin
    TFMul(Ecran).Caption :='Salariés non affectés aux ressources';
    UpdateCaption(Ecran);
    if (Q_Mul <> NIL) Then
      if (TFMul(Ecran).DBListe  <>'PGMULSALARIE') then TFMul(Ecran).SetDBListe('PGMULSALARIE');
    SetControlText('XX_WHERE','PSA_SALARIE NOT IN (SELECT ARS_SALARIE FROM RESSOURCE WHERE ARS_TYPERESSOURCE="SAL" AND ARS_SALARIE<>"")');
    End
  else
    Begin
    TFMul(Ecran).Caption :='Salariés affectés aux ressources';
    UpdateCaption(Ecran);
    if (Q_Mul <> NIL) Then
      if (TFMul(Ecran).DBListe <>'PGMULRESSAFFECT') then TFMul(Ecran).SetDBListe('PGMULRESSAFFECT');
    End;
end;

procedure TOF_PGSalarie_res.ExitEdit(Sender: TObject);
var edit : thedit;
begin
  edit:=THEdit(Sender);
  if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
       edit.text:=AffectDefautCode(edit,10);
end;

procedure TOF_PGSalarie_res.GrilleDblClick(Sender: TObject);
Var Salarie : String;
begin
  {$IFDEF EAGLCLIENT}
  TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row-1) ;
  if TFMul(Ecran).Fetchlestous then
     TheMulQ := TOB(Ecran.FindComponent('Q'))
  else begin PgiBox ('Vous n''avez pas de ligne total dans votre liste, #13#10 Traitement impossible ', Ecran.Caption);exit; end;
  {$ELSE}
  TheMulQ := THQuery(Ecran.FindComponent('Q'));
  {$ENDIF}

  Salarie:=Q_Mul.FindField('PSA_SALARIE').AsString;
  if (TFMul(Ecran).DBListe = 'PGMULSALARIE') then
    Begin
    AglLanceFiche('PAY','RES_SALARIE','','','CREATION;'+Salarie+';;'+Q_Mul.Liste);
    BtnCherche.click;
    End
  Else
    if Q_Mul.Liste  = 'PGMULRESSAFFECT' Then
      Begin
      AglLanceFiche('PAY','RES_SALARIE','','','MODIFICATION;'+Salarie+';'+Q_Mul.FindField('ARS_RESSOURCE').AsString+';'+Q_Mul.Liste);
      BtnCherche.click;
      End;
  TheMulQ:=Nil;
end;

procedure TOF_PGSalarie_res.OnClickAffectAuto(Sender: TObject);
begin    { Création automatique des ressources salariés }
  If PGIAsk('Voulez-vous effectuer une création automatique des ressources?',Ecran.Caption)= Mrno then exit;
  If Q_Mul <> nil then
    if not Q_Mul.eof then
      Begin
      Q_Mul.First;
      InitMoveProgressForm (NIL,'Création automatique des ressources.', 'Veuillez patienter SVP ...',Q_Mul.RecordCount,TRUE,TRUE);
      While Not Q_Mul.Eof do
        Begin
        MoveCurProgressForm('Salarié : '+Q_Mul.FindField('PSA_SALARIE').AsString);
        PGCreeRessource(Q_Mul.FindField('PSA_SALARIE').AsString,'',True);
        Q_Mul.next;
        End;
      FiniMoveProgressForm;
      PgiInfo('Traitement terminé.',Ecran.Caption);
      BtnCherche.click;
      End;
end;


{ TOF_PGRessource : Fiche de saisie Ressource type ="SAL"}

procedure TOF_PGRessource.ChangeRess(Sender: TObject);
begin
Modifier := True;
end;

procedure TOF_PGRessource.ChangeSal(Sender: TObject);
begin
Modifier := True;
end;

procedure TOF_PGRessource.ExitEdit(Sender: TObject);
var edit : thedit;
    Q    : TQuery;
begin
  if Supprimer then exit; { PT1 }
  edit:=THEdit(Sender);
  if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
       edit.text:=AffectDefautCode(edit,10);
  MajCaption;
  EffaceZoneSalarie;
  if (Trim(Edit.text)<>'') then
    Begin
    SetControlText('ARS_TYPERESSOURCE' , 'SAL');
    Q:=OpenSql('SELECT PSA_ETABLISSEMENT,PSA_LIBELLE,PSA_PRENOM,PSA_DATESORTIE,'+
    'PSA_ADRESSE1,PSA_ADRESSE2,PSA_ADRESSE3,PSA_CODEPOSTAL,PSA_VILLE,PSA_PAYS,'+
    'PSA_TELEPHONE,PSA_PORTABLE,PSA_NUMEROSS '+
    'FROM SALARIES WHERE PSA_SALARIE="'+Edit.text+'"',True);
    if Not Q.eof then
      Begin
      SetControlText('ARS_ETABLISSEMENT' , Q.FindField('PSA_ETABLISSEMENT').Asstring);
      SetControlText('ARS_LIBELLE'       , Q.FindField('PSA_LIBELLE').Asstring);
      SetControlText('ARS_LIBELLE2'      , Q.FindField('PSA_PRENOM').Asstring);
      if Q.FindField('PSA_DATESORTIE').AsDateTime<>idate1900 then
        SetControlText('ARS_DATESORTIE'    , DateToStr(Q.FindField('PSA_DATESORTIE').AsDateTime));
      SetControlText('ARS_ADRESSE1'      , Q.FindField('PSA_ADRESSE1').Asstring);
      SetControlText('ARS_ADRESSE2'      , Q.FindField('PSA_ADRESSE2').Asstring);
      SetControlText('ARS_ADRESSE3'      , Q.FindField('PSA_ADRESSE3').Asstring);
      SetControlText('ARS_CODEPOSTAL'    , Q.FindField('PSA_CODEPOSTAL').Asstring);
      SetControlText('ARS_VILLE'         , Q.FindField('PSA_VILLE').Asstring);
      SetControlText('ARS_PAYS'          , Q.FindField('PSA_PAYS').Asstring);
      SetControlText('ARS_TELEPHONE'     , Q.FindField('PSA_TELEPHONE').Asstring);
      SetControlText('ARS_TELEPHONE2'    , Q.FindField('PSA_PORTABLE').Asstring);
      SetControlText('ARS_IMMAT'         , Q.FindField('PSA_NUMEROSS').Asstring);
      End;
    Ferme(Q);
    End
  else { DEB PT1 }
    if (GetControlText('ARS_RESSOURCE')<>'') then
    Begin
    SetControlText('ARS_TYPERESSOURCE' , 'SAL');
    Q:=OpenSql('SELECT ARS_ETABLISSEMENT,ARS_LIBELLE,ARS_LIBELLE2,ARS_DATESORTIE,'+
    'ARS_ADRESSE1,ARS_ADRESSE2,ARS_ADRESSE3,ARS_CODEPOSTAL,ARS_VILLE,ARS_PAYS,'+
    'ARS_TELEPHONE,ARS_TELEPHONE2,ARS_IMMAT '+
    'FROM RESSOURCE WHERE ARS_RESSOURCE="'+GetControlText('ARS_RESSOURCE')+'"',True);
    if Not Q.eof then
      Begin
      SetControlText('ARS_ETABLISSEMENT' , Q.FindField('ARS_ETABLISSEMENT').Asstring);
      SetControlText('ARS_LIBELLE'       , Q.FindField('ARS_LIBELLE').Asstring);
      SetControlText('ARS_LIBELLE2'      , Q.FindField('ARS_LIBELLE2').Asstring);
      if Q.FindField('ARS_DATESORTIE').AsDateTime<>idate1900 then
        SetControlText('ARS_DATESORTIE'    , DateToStr(Q.FindField('ARS_LIBELLE2').AsDateTime));
      SetControlText('ARS_ADRESSE1'      , Q.FindField('ARS_ADRESSE1').Asstring);
      SetControlText('ARS_ADRESSE2'      , Q.FindField('ARS_ADRESSE2').Asstring);
      SetControlText('ARS_ADRESSE3'      , Q.FindField('ARS_ADRESSE3').Asstring);
      SetControlText('ARS_CODEPOSTAL'    , Q.FindField('ARS_CODEPOSTAL').Asstring);
      SetControlText('ARS_VILLE'         , Q.FindField('ARS_VILLE').Asstring);
      SetControlText('ARS_PAYS'          , Q.FindField('ARS_PAYS').Asstring);
      SetControlText('ARS_TELEPHONE'     , Q.FindField('ARS_TELEPHONE').Asstring);
      SetControlText('ARS_TELEPHONE2'    , Q.FindField('ARS_TELEPHONE2').Asstring);
      SetControlText('ARS_IMMAT'         , Q.FindField('ARS_IMMAT').Asstring);
      End;
    Ferme(Q);
    End;  { FIN PT1 }
end;

procedure TOF_PGRessource.MajCaption;
begin
Ecran.caption:='Ressource de : '+GetControlText('ARS_SALARIE')+' '+GetControlText('ARS_LIBELLE');
UpdateCaption(Ecran);
end;

procedure TOF_PGRessource.OnArgument(stArgument: String);
var
  Salarie,Ressource : string;
 {$IFNDEF EAGLCLIENT}
  Btn               : TToolbarButton97;
 {$ENDIF}
begin
  inherited;
  Modifier    := False;
  Mode        := ReadTokenSt(stArgument);
  Salarie     := ReadTokenSt(stArgument);
  Ressource   := ReadTokenSt(stArgument);
  if Trim(stArgument)<>'' then  Liste := ReadTokenSt(stArgument);
  SetControlText('ARS_SALARIE'  , Salarie);
  SetControlText('ARS_RESSOURCE', Ressource);
  THEDit(GetControl('ARS_SALARIE')).OnExit:=ExitEdit;
  THEDit(GetControl('ARS_SALARIE')).OnChange:=ChangeSal;
  THEDit(GetControl('ARS_RESSOURCE')).OnChange:=ChangeRess;
  {$IFNDEF EAGLCLIENT}
  Btn := TToolbarButton97(GetControl('BFirst'));
  if Btn <> nil then Btn.OnClick:=ClickNavigation;
  Btn := TToolbarButton97(GetControl('BPrev'));
  if Btn <> nil then Btn.OnClick:=ClickNavigation;
  Btn := TToolbarButton97(GetControl('BNext'));
  if Btn <> nil then Btn.OnClick:=ClickNavigation;
  Btn := TToolbarButton97(GetControl('Blast'));
  if Btn <> nil then Btn.OnClick:=ClickNavigation;
  {$ELSE}
  SetControlVisible('BFirst',False);
  SetControlVisible('BPrev',False);
  SetControlVisible('BNext',False);
  SetControlVisible('Blast',False);
  {$ENDIF}

  {$IFDEF EAGLCLIENT}
   QMUL:=THQuery(TFVierge(Ecran).FMULQ).TQ ;
  {$ELSE}
   QMUL:=TFVierge(Ecran).FMULQ ;
  {$ENDIF}
end;


procedure TOF_PGRessource.OnLoad;
begin
  inherited;
  Supprimer   := False; { PT1 }
  ExitEdit(THEDit(GetControl('ARS_SALARIE')));
  SetControlEnabled('ARS_RESSOURCE',(Mode='CREATION'));
  SetControlEnabled('ARS_SALARIE'  ,(VH_Paie.PgTypeAffectRes=''));
  SetControlEnabled('ARS_RESSOURCE',(VH_Paie.PgTypeAffectRes=''));
  if (Mode='CREATION') then
    Begin
    SetControlText('ARS_RESSOURCE',PGGetRessource(GetControlText('ARS_SALARIE'),GetControlText('ARS_LIBELLE'),GetControlText('ARS_LIBELLE2')));
    Modifier    := True;
    end;
  SetControlEnabled('BDelete',(Mode='MODIFICATION') and (GetControlText('ARS_SALARIE')<>''));
  MajCaption;
end;



procedure TOF_PGRessource.OnClose;
begin
  inherited;
if Modifier then
  Begin
  MajCaption;
  if PgiAsk('Voulez-vous enregistrer les modifications?',Ecran.Caption) = MrYes then
     OnUpdate;
  end;
end;

procedure TOF_PGRessource.OnDelete;
Var Ressource : string;
begin
  inherited;
  Ressource := Trim(GetControlText('ARS_RESSOURCE'));
  if Ressource <> '' then
    if PgiAsk('Etes-vous sûr de vouloir dissocier le salarié à la ressource : '+Ressource+' ?',Ecran.Caption) = MrYes then { PT1 }
       Begin
       Supprimer:=True;
       SetControlText('ARS_SALARIE', '');
       //EffaceZoneSalarie;  PT1
       End;
end;




procedure TOF_PGRessource.OnUpdate;
Var
  Salarie,Ressource : String;
begin
  inherited;
  LastError:=0;
  if not Modifier then exit;
  ExitEdit(THEDit(GetControl('ARS_SALARIE')));
  Salarie:=Trim(GetControlText('ARS_SALARIE'));
  Ressource:=Trim(GetControlText('ARS_RESSOURCE'));
  if Not Supprimer then
    Begin
    //Test si salarié exist
    if (Salarie='') Or (ExisteSql('SELECT PSA_SALARIE FROM SALARIES WHERE PSA_SALARIE="'+salarie+'"')=False) then
      LastError:=1;
    //Test validité ressource
    if (Ressource='') then LastError:=2
    else
      if (VH_Paie.PgTypeAffectRes='MAT') and (salarie<>ressource) then LastError:=5;
    if (Mode='CREATION') and (LastError=0) then
      Begin
      //Test si ressource non crée en mode création
      if (ExisteSql('SELECT ARS_RESSOURCE FROM RESSOURCE WHERE ARS_RESSOURCE="'+Ressource+'"')) then
         LastError:=3;
      End
    else   //Test si salarié non affecté en mode modification
      if (Mode='MODIFICATION') and (LastError=0) then
        if (ExisteSql('SELECT ARS_SALARIE FROM RESSOURCE WHERE ARS_SALARIE="'+salarie+'" '+
                     'AND ARS_RESSOURCE<>"'+Ressource+'"')) then
          LastError:=4;
    End;
  if LastError<>0 then
    PgiInfo(TexteMessage[LastError],Ecran.caption)
  else
    Begin
    if Supprimer then
      Begin
      if PGSupprRessource(GetControlText('ARS_RESSOURCE')) then
         Begin
         Modifier := False;
         Supprimer := False ; { PT1 }
         SetControlEnabled('BDelete',False);
         End;
      End
    else
      if PGCreeRessource(Salarie,GetControlText('ARS_RESSOURCE'),True) then
        Begin
        Modifier := False;
        Mode:='MODIFICATION';
        SetControlEnabled('ARS_RESSOURCE',False);
        SetControlEnabled('BDelete',True);
        end
      else
        PgiBox('Une erreur est survenue lors de la validation de la ressource.',Ecran.caption);
    End;
end;
{$IFNDEF EAGLCLIENT}
procedure TOF_PGRessource.ClickNavigation(Sender: TObject);
var Init : word;
    Q    : TQuery;
begin
  if QMul=nil then exit;
  if (Mode='CREATION') OR ((Mode='MODIFICATION') and (supprimer)) then
     Begin
     Init := HShowMessage('1;'+Ecran.caption+';Voulez-vous enregistrez les modifications ?;Q;YN;N;N;','','');
     if Init = mryes then
        Begin
        OnUpdate;
        if LastError<>0 then exit;
        End
     else
        if Init = mrCancel then exit;
     End;

       if TToolBarButton97(Sender).Name='BFirst' then QMul.First
  else if TToolBarButton97(Sender).Name='BPrev'  then QMul.prior
  else if TToolBarButton97(Sender).Name='BNext'  then QMul.Next
  else if TToolBarButton97(Sender).Name='BLast'  then QMul.Last;

  if (Liste = 'PGMULSALARIE') then
    Begin
    Q:=OpenSql('SELECT ARS_RESSOURCE,ARS_SALARIE FROM RESSOURCE '+
     'WHERE ARS_SALARIE="'+QMul.FindField('PSA_SALARIE').AsString+'"',True);
    if not Q.Eof then
      OnArgument('MODIFICATION;'+Q.FindField('ARS_SALARIE').AsString+';'+Q.FindField('ARS_RESSOURCE').AsString)
    else
      OnArgument('CREATION;'+QMul.FindField('PSA_SALARIE').AsString);
    Ferme(Q);
    End
  Else
    if (Liste  = 'PGMULRESSAFFECT') then
       OnArgument('MODIFICATION;'+QMul.FindField('PSA_SALARIE').AsString+';'+QMul.FindField('ARS_RESSOURCE').AsString);
  Onload;
  SetControlEnabled('BFirst', (Not QMul.BOF));
  SetControlEnabled('BPrev' , (Not QMul.BOF));
  SetControlEnabled('BNext' , (Not QMul.EOF));
  SetControlEnabled('BLast' , (Not QMul.EOF));
end;
{$ENDIF}
procedure TOF_PGRessource.EffaceZoneSalarie;
begin
  SetControlText('ARS_ETABLISSEMENT' , '');
  SetControlText('ARS_LIBELLE'       , '');
  SetControlText('ARS_LIBELLE2'      , '');
  SetControlText('ARS_DATESORTIE'    , '');
  SetControlText('ARS_ADRESSE1'      , '');
  SetControlText('ARS_ADRESSE2'      , '');
  SetControlText('ARS_ADRESSE3'      , '');
  SetControlText('ARS_CODEPOSTAL'    , '');
  SetControlText('ARS_VILLE'         , '');
  SetControlText('ARS_PAYS'          , '');
  SetControlText('ARS_TELEPHONE'     , '');
  SetControlText('ARS_TELEPHONE2'    , '');
  SetControlText('ARS_IMMAT'         , '');
end;

Initialization
registerclasses([TOF_PGSalarie_res,TOF_PGRessource]) ;
end.

