{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 11/09/2001
Modifié le ... :   /  /
Description .. : TOM gestion des taux AT par établissement
Mots clefs ... : PAIE
*****************************************************************
PT1   : 19/10/2001 SB V563 Fiche de bug 340 Test si themax<>''
PT2   : 02/05/2002 Ph V582 Controle taux renseigné, alerte si code risque ou
                           section non renseignés
PT3-1 : 11/03/2003 VG V_42 Duplication du taux AT - FQ N°10554
PT3-2 : 11/03/2003 VG V_42 Contrôle de la section et du taux
PT4   : 16/09/2003 PH V_42 mauvais dimensionnement des tableaux pour appel
                           fonction presencecomplexe FQ110
PT5   : 27/02/2004 VG V_50 Le code risque AT doit être en majuscule - FQ N°11081
PT6   : 08/11/2004 PH V_60 Rechargement du cache serevur en CWAS - FQ N°11746
PT7   : 18/11/2004 VG V_60 Ne pas permettre de valider un code section AT si sa
                           valeur n'est pas comprise entre 01 et 99 - FQ N°11782
PT8   : 02/12/2004 VG V_60 En création, on avait un message d'erreur
                           FQ N°11822
PT9   : 02/10/2007 VG V_80 Section AT à 98 interdite
}
unit UTOMTauxAt;

interface
uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls, Spin,
{$IFNDEF EAGLCLIENT}
     db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB, DBCtrls, Fe_Main, Fiche, FichList,
{$ELSE}
     MaineAgl, eFiche, eFichList,
{$ENDIF}
     HCtrls, HEnt1, HMsgBox, UTOM, UTOB, HTB97, PgOutils, PgOutils2;

type
  TOM_TAUXAT = class(TOM)
    procedure OnArgument(stArgument: string); override;
    procedure OnChangeField(F: TField); override;
    procedure OnUpdateRecord; override;
    procedure OnLoadRecord; override;
    procedure OnNewRecord; override;
  private
    ETAB, mode: string;
    procedure ExitDate(Sender: TObject);
    procedure DupliquerTauxAT(Sender: TObject);
  end;



implementation
{ TOM_TAUXAT }

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Philippe DUMET
Créé le ...... : 11/09/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE
*****************************************************************}
procedure TOM_TAUXAT.ExitDate(Sender: TObject);
begin
if GetField('PAT_DATEVALIDITE') = IDate1900 then
   SetField('PAT_DATEVALIDITE', Date);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Philippe DUMET
Créé le ...... : 11/09/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE
*****************************************************************}
procedure TOM_TAUXAT.OnArgument(stArgument: string);
var
{$IFNDEF EAGLCLIENT}
DateValid: THDBEdit;
{$ELSE}
DateValid: THEdit;
{$ENDIF}
BDupliquer: TToolBarButton97;
begin
inherited;

{$IFNDEF EAGLCLIENT}
DateValid := THDBEdit(GetControl('PAT_DATEVALIDITE'));
{$ELSE}
DateValid := THEdit(GetControl('PAT_DATEVALIDITE'));
{$ENDIF}

if DateValid <> nil then
   DateValid.OnExit := ExitDate;
SetControlEnabled('PAT_ETABLISSEMENT', FALSE);

BDupliquer := TToolBarButton97(GetControl('BDUPLIQUER'));
if BDupliquer <> nil then
   BDupliquer.OnClick := DupliquerTauxAT;

if Ecran is TFFicheListe then
   begin
{$IFNDEF EAGLCLIENT}
   TFFicheListe(Ecran).FListe.Columns[0].Visible := FALSE;
{$ELSE}
   TFFicheListe(Ecran).FListe.ColWidths[0] := 0;
{$ENDIF}
   end;
ETAB := Trim(StArgument);
ETAB := ReadTokenSt(ETAB); // Recup Etablissement sur lequel on travaille
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Philippe DUMET
Créé le ...... : 11/09/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE
*****************************************************************}
procedure TOM_TAUXAT.OnChangeField(F: TField);
var
BNew: TToolBarButton97;
begin
inherited;
BNew := TToolBarButton97(GetControl('BINSERT'));

SetControlEnabled('BDUPLIQUER', BNew.Enabled);

if (F.FieldName = ('PAT_TAUXAT')) then
   begin
   if ((GetField('PAT_TAUXAT') <> 0) and (GetField('PAT_TAUXAT') >= 100)) then
      begin
      PGIBox('La valeur saisie est incohérente', 'Taux AT');
      SetField('PAT_TAUXAT', 0);
      SetFocuscontrol('PAT_TAUXAT');
      end;
   end;

if (F.FieldName = ('PAT_SECTIONAT')) then
   begin
{PT9
   if ((GetField ('PAT_SECTIONAT')<'01') or
      (GetField ('PAT_SECTIONAT')>'99')) then
}
   if ((GetField ('PAT_SECTIONAT')<'01') or
      (GetField ('PAT_SECTIONAT')>'99') or
      (GetField ('PAT_SECTIONAT')='98')) then
//FIN PT9
      begin
      PGIBox('La valeur saisie est incohérente', 'Section AT');
      SetField('PAT_SECTIONAT', '01');
      SetFocuscontrol('PAT_SECTIONAT');
      end;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Philippe DUMET
Créé le ...... : 11/09/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE
*****************************************************************}
procedure TOM_TAUXAT.OnLoadRecord;
begin
inherited;
SetControlText('ORDREAT', GetField('PAT_ORDREAT'));
if (DS.State in [dsInsert]) and (Mode <> 'DUPLICATION') then
   SetControlEnabled('ORDREAT', TRUE)
else
   SetControlEnabled('ORDREAT', FALSE);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Philippe DUMET
Créé le ...... : 11/09/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE
*****************************************************************}
procedure TOM_TAUXAT.OnNewRecord;
var
QQ: TQuery;
TheMax: string;
begin
inherited;
QQ:= OpenSQL ('SELECT MAX(PAT_ORDREAT) FROM TAUXAT WHERE'+
              ' PAT_ETABLISSEMENT="' + ETAB + '"', TRUE);
if not QQ.EOF then
   begin
   TheMax := QQ.Fields[0].AsString;
   if TheMax <> '' then
      TheMax := IntToStr(StrToInt(TheMax) + 1)
   else TheMax := '1';
   end
else
   TheMax := '1';
Ferme(QQ);
SetField('PAT_ETABLISSEMENT', ETAB);
SetControlEnabled('PAT_ETABLISSEMENT', FALSE);
if (mode <> 'DUPLICATION') then
   begin
   SetField('PAT_ORDREAT', TheMax);
   SetControlText('ORDREAT', TheMax);
   end;
SetControlEnabled('BDUPLIQUER', False);
SetField('PAT_DATEVALIDITE', Date);
SetField('PAT_SECTIONAT', '01');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Philippe DUMET
Créé le ...... : 11/09/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE
*****************************************************************}
procedure TOM_TAUXAT.OnUpdateRecord;
begin
inherited;

if (DS.State in [dsInsert]) then
   begin // Uniquement en creation
   SetField('PAT_ORDREAT', GetControlText('ORDREAT'));
   SetField('PAT_ETABLISSEMENT', ETAB);
   end;

SetField('PAT_CODERISQUE', PGUpperCase(GetField('PAT_CODERISQUE')));
//Controle taux renseigné, alerte si code risque ou section non renseignés
if GetField('PAT_CODERISQUE') = '' then
   begin
   LastError := 1;
   LastErrorMsg := 'Attention, vous n''avez pas renseigné le code risque';
   SetFocusControl('PAT_CODERISQUE');
   end
else
   if (Length(GetField('PAT_SECTIONAT'))) <> 2 then
      begin
      LastError := 1;
      LastErrorMsg := 'Attention, vous n''avez pas correctement renseigné la section AT';
      SetFocusControl('PAT_SECTIONAT');
      end
   else
      if GetField('PAT_TAUXAT') = 0 then
         begin
         LastError := 1;
         LastErrorMsg := 'Vous devez renseigner le taux';
         SetFocusControl('PAT_TAUXAT');
         end;

if (LastError <> 1) then
   begin
   SetControlEnabled('BDUPLIQUER', True);
{$IFDEF EAGLCLIENT}
   AvertirCacheServer('TAUXAT');
{$ENDIF}
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 11/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE
*****************************************************************}
procedure TOM_TAUXAT.DupliquerTauxAT(Sender: TObject);
var
AncBureau, AncCode, AncEtab, AncLibelle, AncRisque, AncSection: string;
AncTauxAT: string;
//Mauvais dimensionnement des tableaux pour appel fonction presencecomplexe
Champ: array[1..3] of Hstring;
Valeur: array[1..3] of variant;
Ok: Boolean;
begin
SetControlEnabled('BDUPLIQUER', False);
//TFFiche(Ecran).BValider.Click;
AncEtab := GetControlText('PAT_ETABLISSEMENT');
AncCode := GetControlText('PAT_ORDREAT');
AncLibelle := GetControlText('PAT_LIBELLE');
AncRisque := GetControlText('PAT_CODERISQUE');
AncTauxAT := GetControlText('PAT_TAUXAT');
AncSection := GetControlText('PAT_SECTIONAT');
AncBureau := GetControlText('PAT_CODEBUREAU');
mode := 'DUPLICATION';
Champ[1] := 'PAT_ETABLISSEMENT';
Valeur[1] := AncEtab;
Champ[2] := 'PAT_ORDREAT';
Valeur[2] := AncCode;
Champ[3] := 'PAT_DATEVALIDITE';
//Date sous forme SQL
Valeur[3] := UsDateTime(Date);
Ok := RechEnrAssocier('TAUXAT', Champ, Valeur);
if Ok = False then //Test si code existe ou non
   begin
   TFFiche(Ecran).Binsert.Click;
   SetField('PAT_ETABLISSEMENT', AncEtab);
   SetControlText('ORDREAT', AncCode);
   SetField('PAT_ORDREAT', AncCode);
   SetField('PAT_LIBELLE', AncLibelle);
   SetField('PAT_CODERISQUE', AncRisque);
   SetField('PAT_TAUXAT', AncTauxAT);
   SetField('PAT_SECTIONAT', AncSection);
   SetField('PAT_CODEBUREAU', AncBureau);
   end
else
   HShowMessage('5;Taux AT :;La duplication est impossible, l''élément existe déjà.;W;O;O;O;;;', '', '');
mode := '';
end;

initialization
  registerclasses([TOM_TAUXAT]);
end.

