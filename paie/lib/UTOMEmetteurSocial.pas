{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 22/11/2001
Modifié le ... :   /  /
Description .. : TOM DADS-U Emetteur Social
Mots clefs ... : PAIE
*****************************************************************}
{
PT1   : 19/08/2002 VG V585 On ne permet pas de valider la fiche si on gère
                           l'accusé réception et média et/ou indicatif n'est pas
                           renseigné
        Refonte complète du source avec suppression des dbcontrols
PT2   : 16/09/2003 VG V_42 Adaptation cahier des charges V7R01
PT3   : 26/02/2004 VG V_50 Contrôle de l'indicatif d'envoi uniquement si Média
                           est "E-Mail" - FQ N°11079
PT4   : 28/12/2004 JL V_60 11086 Contrôle raison sociale
                   + PT4-1 07/02/2005 Ajout verif autre caractère que espace
PT5   : 28/12/2004 VG V_60 Griser l'indicatif si diseable - FQ N°11644
PT6-1 : 22/03/2006 VG V_65 Ajout de nouvelles zones émetteur + réécriture de
                           OnUpdateRecord
PT6-2 : 23/03/2006 VG V_65 Renforcement des contrôles DADS-U
PT7   : 09/10/2006 VG V_70 Modification des contrôles DADS-U - procédure
                           générique
PT8   : 24/11/2006 VG V_70 On force les contrôles, même si pas de changement
                           FQ N°13704
PT9   : 25/04/2007 FC V_72 FQ 12989 Permettre la duplication d'un émetteur
                           social
PT10  : 03/04/2007 MF V_80 Mise en place des WebServices Jedeclare - uniquement dans un environnement PCL
PT11  : 29/08/2007 VG V_80 Adaptation cahier des charges V8R05
PT12  : 05/12/2007 VG V_80 Affichage du caractère interdit - FQ N°14961
}
unit UTOMEmetteurSocial;

interface

uses  StdCtrls,
      Controls,
      Classes,
      forms,
      sysutils,
      ComCtrls,
{$IFNDEF EAGLCLIENT}
      db,
      Fiche, HDB,Fe_Main,DBCtrls,
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
      eFiche, MaineAgl,UtileAGL,
{$ENDIF}
      HCtrls,
      HEnt1,
      HMsgBox,
      UTOM,
      PgOutils,
      PgOutils2,
      PGDADSControles,
      UTOB,
      HTB97 ;

Type
     TOM_EMETTEURSOCIAL  =  Class(TOM)
       procedure OnNewRecord ; override ;
       procedure OnChangeField (F : TField)  ; override ;
       procedure OnArgument (stArgument : String ) ; override ;
       procedure OnLoadRecord ; override ;
       procedure OnUpdateRecord ; override;
       procedure OnClose ; override;

       private
         Mode : String;
         procedure DupliquerEmetteur(Sender: TObject);    //PT9
         procedure LanceComptesJDC(Sender: TObject);
{PT7
       function ControleTableCaractere(Donnee, Champ, Nom : string;
                Identite : boolean) : String;
}
     END ;

implementation
var
NomChamp : array[0..18] of string= ('PET_RAISONSOC',
                                    'PET_COMPLADR',
                                    'PET_COMPLADR2',
                                    'PET_NUMEROVOIE',
                                    'PET_NOMVOIE',
                                    'PET_INSEECOMMUNE',
                                    'PET_VILLE',
                                    'PET_CODEPOSTAL',
                                    'PET_BURDISTRIB',
                                    'PET_CREINDICATIF',
                                    'PET_APPEL1DUDS',
                                    'PET_TEL1DADSU',
                                    'PET_FAX1DADSU',
                                    'PET_APPEL2DUDS',
                                    'PET_TEL2DADSU',
                                    'PET_FAX2DADSU',
                                    'PET_APPEL3DUDS',
                                    'PET_TEL3DADSU',
                                    'PET_FAX3DADSU');
Segment : array[0..18] of string= ('S10.G01.00.002',
                                   'S10.G01.00.003.001',
                                   'S10.G01.00.003.001',
                                   'S10.G01.00.003.003',
                                   'S10.G01.00.003.006',
                                   'S10.G01.00.003.007',
                                   'S10.G01.00.003.009',
                                   'S10.G01.00.003.010',
                                   'S10.G01.00.003.012',
                                   'S10.G01.00.015.001',
                                   'S10.G01.01.005',
                                   'S10.G01.01.006',
                                   'S10.G01.01.007',
                                   'S10.G01.01.005',
                                   'S10.G01.01.006',
                                   'S10.G01.01.007',
                                   'S10.G01.01.005',
                                   'S10.G01.01.006',
                                   'S10.G01.01.007');


{***********A.G.L.Privé.*****************************************
Auteur  ...... : JL
Créé le ...... : 22/11/2001
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_EMETTEURSOCIAL.OnNewRecord ;
begin
Inherited ;
SetField('PET_CREDADSU','-');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : JL
Créé le ...... : 22/11/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_EMETTEURSOCIAL.OnChangeField (F : TField)  ;
begin
Inherited;
if (F.FieldName='PET_CREDADSU') then
   begin
{PT11
   If (GetField ('PET_CREDADSU')='X') Then
      SetControlEnabled ('PET_CREMEDIA', True)
   Else
      begin
      SetField ('PET_CREMEDIA', '');
      SetControlEnabled ('PET_CREMEDIA', False);
      end;
}
   If (GetField ('PET_CREDADSU')='X') Then
      begin
      SetControlEnabled ('TPET_CREINDICATIF', True);
      SetControlEnabled ('PET_CREINDICATIF', True);
      end
   else
      begin
      SetControlEnabled ('TPET_CREINDICATIF', False);
      SetControlEnabled ('PET_CREINDICATIF', False);
      end;
//FIN PT11
   end;

{PT11
if (F.FieldName = 'PET_CREMEDIA') then
   begin
   If GetField('PET_CREMEDIA') = '03' Then
      begin
      SetControlEnabled('TPET_CREINDICATIF',True);
      SetControlEnabled('PET_CREINDICATIF',True);
      SetControlEnabled('TPET_CRECIVILITE', False);
      SetControlEnabled('PET_CRECIVILITE', False);
      SetControlEnabled('PET_CRENOM', False);
      end
   else
   If GetField('PET_CREMEDIA') = '05' Then
      begin
      SetControlEnabled('TPET_CREINDICATIF',False);
      SetControlEnabled('PET_CREINDICATIF',False);
      SetControlEnabled('TPET_CRECIVILITE', True);
      SetControlEnabled('PET_CRECIVILITE', True);
      SetControlEnabled('PET_CRENOM', True);
      end
   else
      begin
      SetControlEnabled('TPET_CREINDICATIF',False);
      SetControlEnabled('PET_CREINDICATIF',False);
      SetControlEnabled('TPET_CRECIVILITE', False);
      SetControlEnabled('PET_CRECIVILITE', False);
      SetControlEnabled('PET_CRENOM', False);
      end;
   end;
}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : JL
Créé le ...... : 22/11/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_EMETTEURSOCIAL.OnArgument(stArgument: String);
var
  Btn: TToolBarButton97;
  BtnJDC : TToolBarButton97;       // PT10
begin
  Inherited;

  //DEB PT9
  Mode := '';
  Btn := TToolBarButton97(GetControl('BDUPLIQUER'));
  if btn <> nil then Btn.OnClick := DupliquerEmetteur;
  //FIN PT9
// d PT10
  if (V_PGI.ModePCL='1') then
  begin
    BtnJDC := TToolBarButton97(GetControl('BCOMPTEJDC'));
    SetControlProperty('BCOMPTEJDC', 'Visible', true);
    if BtnJDC <> nil then BtnJDC.OnClick := LanceComptesJDC;
  end;
// f PT10
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : JL
Créé le ...... : 22/11/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_EMETTEURSOCIAL.OnLoadRecord;
begin
Inherited ;
{PT11
If GetField('PET_CREDADSU') = 'X' Then
   SetControlEnabled('PET_CREMEDIA',True)
Else
   SetControlEnabled('PET_CREMEDIA',False);
}   
//PT8
if not (ds.state in [dsinsert,dsedit]) then
   ds.edit;
//FIN PT8
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : JL
Créé le ...... : 22/11/2001
Modifié le ... : 22/03/2006
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_EMETTEURSOCIAL.OnUpdateRecord;
Var
CInterdit, Mes, MessCP, mesINSEE, mesSIRET, Valeur, StLexique : String;
i : integer;
QRechLexique : TQuery;
TLexiqueD : TOB;
begin
Inherited;
//PT7
if (TLexique=nil) then
   begin
   StLexique:= 'SELECT *'+
               ' FROM DADSLEXIQUE WHERE'+
               ' PDL_DADSSEGMENT LIKE "S10%" AND'+
               ' PDL_EXERCICEFIN = ""'+
               ' ORDER BY PDL_DADSSEGMENT';
   QRechLexique:=OpenSql(StLexique, TRUE);
   TLexique:= TOB.Create('Le Lexique DADS-U', NIL, -1);
   TLexique.LoadDetailDB('LEXIQUE','','',QRechLexique,False);
   Ferme(QRechLexique);
   end;

ErrorDADSU:= 0;
i:= 0;
While ((ErrorDADSU=0) and (i<19)) do
      begin
      Valeur:= GetControlText(NomChamp[i]);
      TLexiqueD:= TLexique.FindFirst(['PDL_DADSSEGMENT'],[Segment[i]],FALSE);

{PT12
      ErrorDADSU:= ControleForme (Segment[i], Valeur, TLexiqueD, True);
}
      ErrorDADSU:= ControleForme (Segment[i], Valeur, TLexiqueD, True, CInterdit);
//FIN ggPT12g
      if (ErrorDADSU=7) then
         ErrorDADSU:=0;
      if (GetControlText (NomChamp[i])<>Valeur) then
         SetControlText (NomChamp[i], Valeur);
      i:= i+1;
      end;
If ErrorDADSU <> 0 then
   begin
   LastError:= 1;
{PT12
   Mes:= RenvoieErreur (ErrorDADSU);
}
   Mes:= RenvoieErreur (ErrorDADSU, CInterdit);
//FIN PT12
   PGIBox (Mes+' :#13#10'+Valeur, TLexiqueD.GetValue('PDL_DADSNOM'));
   SetFocusControl(NomChamp[i-1]);
   end;

Mes:= '';
//FIN PT7
mesSIRET:= '';
MessCP:= '';
mesINSEE:= '';

//Zones obligatoires
If (GetControlText('PET_SIRET') = '') Then
   begin
   Mes:= Mes+'#13#10    - le N° SIRET';
   SetFocusControl ('PET_SIRET');
   end
else
   begin
   if ((IsNumeric(GetControlText('PET_SIRET')) = False) or
      (ControlSiret(GetControlText('PET_SIRET')) = False)) Then
      begin
      MesSiret:= '#13#10 - le SIRET n''est pas valide';
      SetFocusControl ('PET_SIRET');
      end;
   end;

If (GetControlText('PET_RAISONSOC') = '') Then
   begin
   Mes:= Mes+'#13#10    - la Raison Sociale';
   SetFocusControl ('PET_RAISONSOC');
   end;

If (GetControlText('PET_CODEPOSTAL') = '') then
   begin
   Mes:= Mes+'#13#10    - le Code Postal';
   SetFocusControl ('PET_CODEPOSTAL');
   end;

If (GetControlText('PET_BURDISTRIB') = '') then
   begin
   Mes:= Mes+'#13#10    - le Bureau Distributeur';
   SetFocusControl ('PET_BURDISTRIB');
   end;

If (GetControlText('PET_CIVIL1DADSU') = '') then
   begin
   Mes:= Mes+'#13#10    - la civilité du 1er contact';
   SetFocusControl ('PET_CIVIL1DADSU');
   end;

If (GetControlText('PET_CONT1DUDS') = '') then
   begin
   Mes:= Mes+'#13#10    - le nom du 1er contact';
   SetFocusControl ('PET_CONT1DUDS');
   end;

If GetControlText ('PET_DOMAINEDUDS1') = '' then
   begin
   Mes:= Mes+'#13#10    - le domaine du 1er contact';
   SetFocusControl ('PET_DOMAINEDUDS1');
   end;

If (GetControlText('PET_APPEL1DUDS') = '') then
   begin
   Mes:= Mes+'#13#10    - l''email du 1er contact';
   SetFocusControl('PET_APPEL1DUDS');
   end;

If (GetControlText('PET_TEL1DADSU') = '') then
   begin
   Mes:= Mes+'#13#10    - le téléphone du 1er contact';
   SetFocusControl('PET_TEL1DADSU');
   end;

//Personne à contacter 2 si une zone remplie
If ((GetControlText ('PET_CIVIL2DADSU') <> '') or
   (GetControlText ('PET_CONT2DUDS') <> '') or
   (GetControlText ('PET_DOMAINEDUDS2') <> '') or
   (GetControlText ('PET_APPEL2DUDS') <> '') or
   (GetControlText ('PET_TEL2DADSU') <> '') or
   (GetControlText ('PET_FAX2DADSU') <> '')) then
   begin
   If (GetControlText('PET_CIVIL2DADSU') = '') then
      begin
      Mes:= Mes+'#13#10    - la civilité du 2ème contact';
      SetFocusControl ('PET_CIVIL2DADSU');
      end;
   If (GetControlText ('PET_CONT2DUDS') = '') Then
      begin
      Mes:= Mes+'#13#10    - le nom du 2ème contact';
      SetFocusControl ('PET_CONT2DUDS');
      end;
   If GetControlText ('PET_DOMAINEDUDS2') = '' Then
      begin
      Mes:= Mes+'#13#10    - le domaine du 2ème contact';
      SetFocusControl ('PET_DOMAINEDUDS2');
      end;
   If (GetControlText ('PET_APPEL2DUDS') = '') Then
      begin
      Mes:= Mes+'#13#10    - l''email du 2ème contact';
      SetFocusControl ('PET_APPEL2DUDS');
      end;
   If (GetControlText('PET_TEL2DADSU') = '') then
      begin
      Mes:= Mes+'#13#10    - le téléphone du 2ème contact';
      SetFocusControl('PET_TEL2DADSU');
      end;
   end;

//Personne à contacter 3 si une zone remplie
If ((GetControlText ('PET_CIVIL3DADSU') <> '') or
   (GetControlText ('PET_CONT3DUDS') <> '') or
   (GetControlText ('PET_DOMAINEDUDS3') <> '') or
   (GetControlText ('PET_APPEL3DUDS') <> '') or
   (GetControlText ('PET_TEL3DADSU') <> '') or
   (GetControlText ('PET_FAX3DADSU') <> '')) then
   begin
   If (GetControlText('PET_CIVIL3DADSU') = '') then
      begin
      Mes:= Mes+'#13#10    - la civilité du 3ème contact';
      SetFocusControl ('PET_CIVIL3DADSU');
      end;
   If (GetControlText ('PET_CONT3DUDS') = '') Then
      begin
      Mes:= Mes+'#13#10    - le nom du 3ème contact';
      SetFocusControl ('PET_CONT3DUDS');
      end;
   If GetControlText ('PET_DOMAINEDUDS3') = '' Then
      begin
      Mes:= Mes+'#13#10    - le domaine du 3ème contact';
      SetFocusControl ('PET_DOMAINEDUDS3');
      end;
   If (GetControlText ('PET_APPEL3DUDS') = '') Then
      begin
      Mes:= Mes+'#13#10    - l''email du 3ème contact';
      SetFocusControl ('PET_APPEL3DUDS');
      end;
   If (GetControlText('PET_TEL3DADSU') = '') then
      begin
      Mes:= Mes+'#13#10    - le téléphone du 3ème contact';
      SetFocusControl('PET_TEL3DADSU');
      end;
   end;

// DUCS NOM
If GetControlText('PET_CONTACTDUCS') = '' then
   begin
   Mes:= Mes+'#13#10    - le nom de l''émetteur de la DUCS';
   SetFocusControl('PET_CONTACTDUCS');
   end;
// ducs tel ou fax
If (GetControlText('PET_TELDUCS') = '') AND
   (GetControlText('PET_FAXDUCS') = '') then
   begin
   Mes:= Mes+'#13#10    - le téléphone ou le fax de l''émetteur de la DUCS';
   SetFocusControl('PET_TELDUCS');
   end;

if ((GetControlText('PET_CODEPOSTAL') <> '') and
   (GetControlText('PET_PAYS')='FRA')) then
   begin
   If Not (IsNumeric(GetControlText('PET_CODEPOSTAL')))  then
      begin
      messCP:= messCP+'#13#10 - le code postal doit être numérique';
      SetFocusControl('PET_CODEPOSTAL');
      end
   Else
   If ValeurI (GetControlText('PET_CODEPOSTAL')) = 0 Then
      begin
      messCP:= messCP+'#13#10 - le code postal ne peut pas avoir la valeur 0';
      SetFocusControl( 'PET_CODEPOSTAL');
      end
   Else
   If length (GetControlText('PET_CODEPOSTAL')) <> 5 Then
      begin
      messCP:= messCP+'#13#10 - le code postal doit comporter 5 caractères';
      SetFocusControl('PET_CODEPOSTAL');
      end;
   end;

If (GetControlText('PET_INSEECOMMUNE') <> '') AND
   (length (GetControlText('PET_INSEECOMMUNE')) <> 5) Then
   begin
   mesINSEE:= mesINSEE+'#13#10 - le code INSEE doit être nul ou comporter 5 caractères';
   SetFocusControl('PET_INSEECOMMUNE');
   end;
if mes <> '' then
   begin
   If (MessCP <> '') or (mesSIRET <> '') or (mesINSEE <> '') Then
      begin
      LastError:= 1;
      LastErrorMsg:= 'Veuillez renseigner :'+mes+
                     '#13#10 #13#10NB : La saisie de ce(s) champ(s) est obligatoire.' +
                     '#13#10 #13#10 Format des champs :'+messCP+mesSIRET+mesINSEE;
      end
   Else
      begin
      LastError:= 1;
      LastErrorMsg:= 'Veuillez renseigner :'+mes+
                     '#13#10 #13#10NB : La saisie de ce(s) champ(s) est obligatoire.';
      end;
   end
Else
If (MessCP <> '') or (mesSIRET <> '') or (mesINSEE <> '') Then
   begin
   LastError:= 1;
   LastErrorMsg:= '#13#10Format des champs :'+messCP+mesSIRET+mesINSEE;
   end;

{PT11
if (mode<>'DUPLICATION') then
   begin
   if ((GetField ('PET_CREDADSU')='X') AND ((GetField ('PET_CREMEDIA')='') OR
      ((GetField ('PET_CREMEDIA')='03') AND (GetField ('PET_CREINDICATIF')='')) or
      ((GetField ('PET_CREMEDIA')='05') AND ((GetField ('PET_CRECIVILITE')='') or
      (GetField ('PET_CRENOM')=''))))) then
      begin
      LastError:= 1;
      LastErrorMsg:= 'Si vous voulez gérer l''accusé réception, vous devez#13#10'+
                     'renseigner un média et les données correspondantes.';
      SetFocusControl ('PET_CREMEDIA');
      end;
}
if ((GetField ('PET_CREDADSU')='X') AND
   (GetField ('PET_CREINDICATIF')='')) then
   begin
   LastError:= 1;
   LastErrorMsg:= 'Si vous voulez gérer l''accusé réception, vous devez#13#10'+
                  'renseigner un email.';
   SetFocusControl ('PET_CREINDICATIF');
   end;
{
   end;
}   
//FIN PT11
end;

//PT7
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 09/10/2006
Modifié le ... :   /  /
Description .. : Sortie de la fiche - libération de la Tob
Mots clefs ... : PAIE
*****************************************************************}
procedure TOM_EMETTEURSOCIAL.OnClose;
begin
FreeAndNil (TLexique);
end;
//FIN PT7


//DEB PT9
procedure TOM_EMETTEURSOCIAL.DupliquerEmetteur(Sender: TObject);
var
{$IFNDEF EAGLCLIENT}
  Code: THDBEdit;
{$ELSE}
  Code: THEdit;
{$ENDIF}
  Champ: array[1..1] of Hstring;
  Valeur: array[1..1] of variant;
  AncValCode : String;
  Ok : Boolean;
begin
  TFFiche(Ecran).BValider.Click;
  AncValCode := GetField('PET_EMETTSOC');
  mode := 'DUPLICATION';
  AglLanceFiche('PAY', 'DUPLI_EMETTEURSOC', '', '', 'ELT;' + AncValCode + ';4');

  if PGCodeDupliquer <> '' then
  begin
    Champ[1] := 'PET_EMETTSOC';
    Valeur[1] := PGCodeDupliquer;
    Ok := RechEnrAssocier('EMETTEURSOCIAL', Champ, Valeur);
    if Ok = False then //Test si code existe ou non
    begin
{$IFNDEF EAGLCLIENT}
      Code := THDBEdit(GetControl('PET_EMETTSOC'));
{$ELSE}
      Code := THEdit(GetControl('PET_EMETTSOC'));
{$ENDIF}
      if (code <> nil) then
        DupliquerPaie(TFFiche(Ecran).TableName, Ecran);
      SetField('PET_EMETTSOC', PGCodeDupliquer);

      TFFiche(Ecran).Bouge(nbPost);
    end
    else
      HShowMessage('5;Emetteur social :;La duplication est impossible, l''émetteur social existe déjà.;W;O;O;O;;;', '', '');
  end;
  mode := '';
end;
//FIN PT9
// d PT10
procedure TOM_EMETTEURSOCIAL.LanceComptesJDC(Sender: TObject);
begin
				AGLLanceFiche ('FIS', 'TDI_COMPTES_JDC', '', '', 'SOCIAL;'+
                        GetControlText('PET_EMETTSOC')+';'+
                        GetControlText('PET_SIRET'));
end;
// f PT10
Initialization
registerclasses([TOM_EMETTEURSOCIAL]) ;
end.
