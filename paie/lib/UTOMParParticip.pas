{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 11/09/2001
Modifié le ... :   /  /
Description .. : TOM gestion des taux AT par établissement
Mots clefs ... : PAIE
*****************************************************************
PT1    : 19/10/2001 SB V563 Fiche de bug 340 Test si themax<>''
PT2    : 02/05/2002 Ph V582 Controle taux renseigné, alerte si code risque ou
                            section non renseignés
PT3-1  : 11/03/2003 VG V_42 Duplication du taux AT - FQ N°10554
PT3-2  : 11/03/2003 VG V_42 Contrôle de la section et du taux
PT4    : 16/09/2003 PH V_421 mauvais dimensionnement des tableaux pour appel fonction presencecomplexe FQ110
}
unit UTOMTauxAt;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,Spin,
{$IFNDEF EAGLCLIENT}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}HDB,DBCtrls,Fe_Main,Fiche,FichList,
{$ELSE}
       MaineAgl,eFiche,eFichList,
{$ENDIF}
       HCtrls,HEnt1,HMsgBox,UTOM,UTOB,HTB97,PgOutils ;

Type
     TOM_TAUXAT = Class(TOM)
       procedure OnArgument (stArgument : String ) ; override ;
       procedure OnChangeField (F : TField)  ; override ;
       procedure OnUpdateRecord  ; override ;
       procedure OnLoadRecord ; override ;
       procedure OnNewRecord  ; override ;
       private
       ETAB,mode : String;
       procedure ExitDate (Sender: TObject);
       procedure DupliquerTauxAT(Sender: TObject);
     END ;



implementation
{ TOM_TAUXAT }

procedure TOM_TAUXAT.ExitDate(Sender: TObject);
begin
if GetField ('PAT_DATEVALIDITE') = IDate1900 then SetField ('PAT_DATEVALIDITE', Date) ;
end;

procedure TOM_TAUXAT.OnArgument(stArgument: String);
var
{$IFNDEF EAGLCLIENT}
  DateValid   : THDBEdit;
{$ELSE}
  DateValid   : THEdit;
{$ENDIF}
BDupliquer:TToolBarButton97;
begin
  inherited;

{$IFNDEF EAGLCLIENT}
  DateValid := THDBEdit (GetControl ('PAT_DATEVALIDITE')) ;
{$ELSE}
  DateValid := THEdit (GetControl ('PAT_DATEVALIDITE')) ;
{$ENDIF}

  if DateValid <> NIL then DateValid.OnExit := ExitDate ;
  SetControlEnabled('PAT_ETABLISSEMENT', FALSE) ;

//PT3-1
BDupliquer:=TToolBarButton97(GetControl('BDUPLIQUER')) ;
if BDupliquer<>nil then
   BDupliquer.OnClick:=DupliquerTauxAT;
//FIN PT3-1

  if Ecran is TFFicheListe then
    BEGIN
{$IFNDEF EAGLCLIENT}
        TFFicheListe(Ecran).FListe.Columns[0].Visible:=FALSE ;
{$ELSE}
        TFFicheListe(Ecran).FListe.ColWidths[0] := 0 ;
{$ENDIF}
    END ;
  ETAB:=Trim (StArgument);
  ETAB:=ReadTokenSt(ETAB);// Recup Etablissement sur lequel on travaille
end;

procedure TOM_TAUXAT.OnChangeField(F: TField);
var
BNew:TToolBarButton97;
begin
  inherited;
BNew:=TToolBarButton97(GetControl('BINSERT')) ;

SetControlEnabled('BDUPLIQUER', BNew.Enabled);

//PT3-2
if (F.FieldName=('PAT_TAUXAT')) then
   begin
   if ((GetField('PAT_TAUXAT') <> 0) and
      (GetField('PAT_TAUXAT') >= 100)) then
      begin
      PGIBox ('La valeur saisie est incohérente','Taux AT');
      SetField('PAT_TAUXAT',0);
      SetFocuscontrol('PAT_TAUXAT');
      end;
   end;
//FIN PT3-2
end;

procedure TOM_TAUXAT.OnLoadRecord;
begin
  inherited;
SetControlText ('ORDREAT', GetField ('PAT_ORDREAT')) ;
{PT3-1
if (DS.State in [dsInsert]) then
}
if (DS.State in [dsInsert]) and (Mode<>'DUPLICATION') then
   SetControlEnabled ('ORDREAT', TRUE)
else
   SetControlEnabled ('ORDREAT', FALSE) ;
end;

procedure TOM_TAUXAT.OnNewRecord;
Var QQ : TQuery ;
    TheMax : String;
begin
  inherited;
  QQ:=OpenSQL('SELECT MAX(PAT_ORDREAT) FROM TAUXAT WHERE PAT_ETABLISSEMENT="'+ETAB+'"',TRUE) ;
  if Not QQ.EOF then
     begin
     TheMax:=QQ.Fields[0].AsString ;
     if TheMax<>'' then       //PT1
       TheMax := IntToStr (StrToInt (TheMax) + 1)
     else TheMax:='1' ;            //PT1
     end  else TheMax:='1' ;
  Ferme(QQ) ;
  SetField ('PAT_ETABLISSEMENT', ETAB) ;
  SetControlEnabled('PAT_ETABLISSEMENT', FALSE) ;
//PT3-1
  if (mode <> 'DUPLICATION') then
     begin
     SetField ('PAT_ORDREAT', TheMax) ;
     SetControlText ('ORDREAT', TheMax) ;
     end;
  SetControlEnabled('BDUPLIQUER', False);
//FIN PT3-1
  SetField ('PAT_DATEVALIDITE', Date) ;
end;

procedure TOM_TAUXAT.OnUpdateRecord;
begin
  inherited;

  if (DS.State in [dsInsert]) then
  begin // Uniquement en creation
    SetField ('PAT_ORDREAT', GetControlText ('ORDREAT')) ;
    SetField ('PAT_ETABLISSEMENT', ETAB) ;
  end;

// PT2 02/05/2002 V582 Ph Controle taux renseigné, alerte si code risque ou section non renseignés
  if GetField ('PAT_CODERISQUE') = '' then
  begin
    LastError := 1 ;
    LastErrorMsg := 'Attention, vous n''avez pas renseigné le code risque' ;
    SetFocusControl ('PAT_CODERISQUE') ;
  end
  else
{PT3-2
  if GetField ('PAT_SECTIONAT') = '' then
}
if (Length(GetField ('PAT_SECTIONAT'))) <> 2 then
    begin
    LastError := 1 ;
    LastErrorMsg := 'Attention, vous n''avez pas correctement renseigné la section AT' ;
    SetFocusControl ('PAT_SECTIONAT') ;
    end
else
    if GetField ('PAT_TAUXAT') = 0 then
       begin
       LastError := 1 ;
       LastErrorMsg:='Vous devez renseigner le taux' ;
       SetFocusControl ('PAT_TAUXAT') ;
       end;
// FIN PT2

//PT3-1
if (LastError <> 1) then
   SetControlEnabled('BDUPLIQUER', True);
//FIN PT3-1

end;


//PT3-1
procedure TOM_TAUXAT.DupliquerTauxAT(Sender: TObject);
var
AncBureau, AncCode, AncEtab, AncLibelle, AncRisque, AncSection : string;
AncTauxAT : string;
// PT4    : 16/09/2003 PH V_421 mauvais dimensionnement des tableaux por appel fonction presencecomplexe
Champ :array[1..3] of string;
Valeur :array[1..3] of variant ;
Ok : Boolean;
begin
SetControlEnabled('BDUPLIQUER', False);
//TFFiche(Ecran).BValider.Click;
AncEtab:=GetControlText('PAT_ETABLISSEMENT');
AncCode:=GetControlText('PAT_ORDREAT');
AncLibelle:=GetControlText('PAT_LIBELLE');
AncRisque:=GetControlText('PAT_CODERISQUE');
AncTauxAT:=GetControlText('PAT_TAUXAT');
AncSection:=GetControlText('PAT_SECTIONAT');
AncBureau:=GetControlText('PAT_CODEBUREAU');
mode:='DUPLICATION';
Champ[1]:= 'PAT_ETABLISSEMENT';
Valeur[1]:= AncEtab;
Champ[2]:= 'PAT_ORDREAT';
Valeur[2]:= AncCode;
Champ[3]:= 'PAT_DATEVALIDITE';
// PT4    : 16/09/2003 PH V_421 Date sous forme SQL
Valeur[3]:= UsDateTime(Date);
Ok:=RechEnrAssocier('TAUXAT',Champ,Valeur);
if Ok=False then         //Test si code existe ou non
   begin
   TFFiche(Ecran).Binsert.Click;
   SetField('PAT_ETABLISSEMENT',AncEtab);
   SetControlText ('ORDREAT', AncCode) ;
   SetField('PAT_ORDREAT',AncCode);
   SetField('PAT_LIBELLE', AncLibelle);
   SetField('PAT_CODERISQUE', AncRisque);
   SetField('PAT_TAUXAT', AncTauxAT);
   SetField('PAT_SECTIONAT', AncSection);
   SetField('PAT_CODEBUREAU', AncBureau);
   end
else
   HShowMessage('5;Taux AT :;La duplication est impossible, l''élément existe déjà.;W;O;O;O;;;','','');
mode:='';

end;
//FIN PT3-1

Initialization
registerclasses([TOM_TAUXAT]) ;
end.
