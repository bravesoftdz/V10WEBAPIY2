{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 06/04/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : TRANSFREF ()
Mots clefs ... : TOF;TRANSFREF
*****************************************************************}
Unit UtofMBOTransfRef ;

Interface

Uses StdCtrls, Controls, Classes, db, forms, sysutils, dbTables, ComCtrls,
     HCtrls, HEnt1, HMsgBox, UTOB, UTOF, M3FP, vierge, PropoAffectTrf,
{$IFDEF EAGLCLIENT}
     Maineagl,
{$ELSE}
     FE_Main,
{$ENDIF}
     AGLInit;

Type
  TOF_MBOTRANSFREF = Class (TOF)
  public
    TobListeEtab, TobMethode, TobAffectationTrf : TOB ;
    CodeMethodeMemorise : string;
    procedure OnArgument (S : String ) ; override ;
    procedure OnLoad                   ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnClose                  ; override ;
    procedure OnChangeCodeArtProp;
    procedure OnClickChoixMethode;
    procedure SetLastError (Num : integer; ou : string );
  end ;

Const
	// libellés des messages
    TexteMessage: array[1..4] of string 	= (
          {1}        'Vous devez renseigner un article',
          {2}        'L''article n''existe pas',
          {3}        'Vous devez renseigner une méthode',
          {4}        'Pas de pièce à transférer pour cet article.'
            );


Implementation

procedure TOF_MBOTRANSFREF.SetLastError (Num : integer; ou : string );
begin
  if ou<>'' then SetFocusControl(ou);
  LastError:=Num;
  LastErrorMsg:=TexteMessage[LastError];
  TForm(Ecran).ModalResult:=0;
end ;

procedure TOF_MBOTRANSFREF.OnArgument (S : String ) ;
var Critere : string ;
begin
  Inherited ;
  Critere:=(ReadTokenSt(S)) ;
  if Critere <> '' then SetControltext('CODPROPAFF', Critere);
end ;

procedure TOF_MBOTRANSFREF.OnLoad ;
var QQ : TQuery ;
begin
  Inherited ;
  TFVierge(Ecran).Top:=312;
  TFVierge(Ecran).Left:=246;

  TobMethode:=TOB.Create('Methode manuelle',Nil,-1) ;

  // Recherche du dépôt correspondant à la proposition
  QQ:=OpenSQL('select GTE_CODEPTRF,GTE_DEPOTEMET from PROPTRANSFENT '+
              'where GTE_CODEPTRF = "'+GetControltext('CODPROPAFF')+'"',True) ;
  if not QQ.EOF then
     begin
     SetControltext('CODDEPOT',QQ.Findfield('GTE_DEPOTEMET').AsString);
     end;
  Ferme(QQ) ;

  // Recherche de la méthode d'affectation de l'utilisateur
  SetControlText('METHODE', '');
  QQ:=OpenSQL('SELECT CPU_GCMCODMETPAFF FROM CPPROFILUSERC WHERE '+
              'CPU_USERGRP="'+V_PGI.Groupe+'" AND CPU_USER="'+V_PGI.User+'" AND '+
              'CPU_TYPE="ETA"', TRUE) ;
  if (Not QQ.EOF) and (QQ.FindField('CPU_GCMCODMETPAFF').AsString<>'') then
     begin
     CodeMethodeMemorise:=QQ.FindField('CPU_GCMCODMETPAFF').AsString ;
     if ExisteSQL('Select * from PROPMETHODE where GTM_CODEMETPAFF="'+CodeMethodeMemorise+'"') then
        SetControlText('METHODE', CodeMethodeMemorise) ;
     end;
  CodeMethodeMemorise:=GetControlText('METHODE');
  Ferme(QQ);
end ;

procedure TOF_MBOTRANSFREF.OnUpdate ;
var TOBProp : TOB;
    TobT1,TobT2,TobT3 : TOB;
    CodeMethode : string;
    CodeDepotEmet,CodePropTrf   : String;
    QQ : TQuery ;
begin
  Inherited ;
  if GetControlText('CODEARTICLE')='' then
     begin
     SetLastError(1, 'CODEARTICLE'); exit ;
     end ;
  if Not ExisteSQL('Select GA_LIBELLE from ARTICLE '+
                   'where GA_CODEARTICLE="'+GetControlText('CODEARTICLE')+'"') then
     begin
     SetLastError(2, 'CODEARTICLE'); exit ;
     end ;

  CodeMethode:=GetControlText('METHODE');
  if GetControlText('METHODE')='' then
     begin
     SetLastError(3, 'METHODE'); exit ;
     end ;

  // Enregistrement de la méthode d'affectation de l'utilisateur
  if CodeMethodeMemorise<>CodeMethode then
     begin
     CodeMethodeMemorise:=CodeMethode;
     if ExisteSQL('SELECT CPU_GCMCODMETPAFF FROM CPPROFILUSERC WHERE '+
                  'CPU_USERGRP="'+V_PGI.Groupe+'" AND CPU_USER="'+V_PGI.User+'" AND '+
                  'CPU_TYPE="ETA"') then
        ExecuteSQL('Update CPPROFILUSERC set CPU_GCMCODMETPAFF="'+CodeMethode+'" '+
                   'where CPU_USERGRP="'+V_PGI.Groupe+'" AND CPU_USER="'+V_PGI.User+'" AND '+
                         'CPU_TYPE="ETA"') else
        ExecuteSQL('Insert Into CPPROFILUSERC (CPU_USERGRP,CPU_USER,CPU_TYPE,CPU_GCMCODMETPAFF) '+
                   'Values ("'+V_PGI.Groupe+'","'+V_PGI.User+'","ETA","'+CodeMethode+'")');
     end;
     
// ----------------------------------------------------------------------------------------
// Lancement du traitement de calcul de la nouvelle affectation de proposition
// de transfert, concernant l'article sélectionné.
// ----------------------------------------------------------------------------------------
TobAffectationTrf:=TOB.Create('Article Selectionne', Nil, -1);
TobT1:=TOB.Create('Article Select', TobAffectationTrf, -1);
TobT1.AddChampSup ('GA_ARTICLE', False);
TobT1.AddChampSup ('GA_CODEARTICLE', False);
TobT1.AddChampSup ('GA_LIBELLE', False);
TobT1.AddChampSup ('GA_STATUTART', False);
TobT1.PutValue    ('GA_CODEARTICLE', GetControlText('CODEARTICLE'));
TobT1.PutValue    ('GA_LIBELLE', GetControlText('DESIGNATION'));
QQ:=OpenSQL('Select GA_ARTICLE,GA_LIBELLE,GA_STATUTART from ARTICLE '+
            'where GA_CODEARTICLE="'+GetControlText('CODEARTICLE')+'" and '+
            'GA_STATUTART<>"DIM"',True);
if Not QQ.EOF then
   begin
   TobT1.PutValue ('GA_ARTICLE', QQ.Findfield('GA_ARTICLE').AsString);
   TobT1.PutValue ('GA_LIBELLE', QQ.Findfield('GA_LIBELLE').AsString);
   TobT1.PutValue ('GA_STATUTART', QQ.Findfield('GA_STATUTART').AsString);
   end;
Ferme(QQ);

if TobMethode.detail.count>0 then
   begin
   if TobMethode.Detail[0].GetValue('CODE')<>CodeMethode then
      TobMethode.ClearDetail ;
   end;
if TobMethode.detail.count=0 then
   begin
   TobT2:=TOB.Create('Methode', TobMethode, -1);
   TobT2.AddChampSup ('ORDRE', False);
   TobT2.AddChampSup ('CODE', False);
   TobT2.AddChampSup ('LIBELLE', False);
   TobT2.AddChampSup ('TYPE', False);
   TobT2.AddChampSup ('UTILCOEF', False);
   TobT2.AddChampSup ('ARRONDI', False);
   TobT2.PutValue    ('ORDRE', '1');
   TobT2.PutValue    ('CODE', CodeMethode);
   TobT2.PutValue    ('UTILCOEF', '-');
   QQ:=OpenSQL('Select GTM_LIBMETPAFF,GTM_TYPEMETPAFF from PROPMETHODE '+
               'where GTM_CODEMETPAFF="'+CodeMethode+'"',True);
   if Not QQ.EOF then
      begin
      TobT2.PutValue ('LIBELLE', QQ.Findfield('GTM_LIBMETPAFF').AsString);
      TobT2.PutValue ('TYPE', QQ.Findfield('GTM_TYPEMETPAFF').AsString);
      end;
   Ferme(QQ);
   end;

TobListeEtab:=TOB.Create('Etablissements destinataires',Nil,-1) ;
QQ:=OpenSQL('Select * from PROPMETETAB where GTQ_CODEMETPAFF="'+
            CodeMethode+'" Order by GTQ_LIGMETPAFF',True) ;
QQ.First;
while Not QQ.EOF do
   begin
   TobT3 := TOB.Create('Etab Dest', TobListeEtab, -1);
   TobT3.AddChampSup('ETABLISSEMENT',False);
   TobT3.AddChampSup('LIBELLE',False);
   TobT3.AddChampSup('COEFF',False);
   TobT3.AddChampSup('POIDS',False);
   TobT3.PutValue('ETABLISSEMENT',QQ.Findfield('GTQ_ETABLISSEMENT').AsString);
   TobT3.PutValue('LIBELLE','');
   TobT3.PutValue('COEFF',QQ.Findfield('GTQ_COEFREPAR').AsFloat);
   TobT3.PutValue('POIDS',QQ.Findfield('GTQ_POIDS').AsInteger);
   QQ.Next;
   end;
Ferme(QQ);
SetFocusControl('CODEARTICLE');

TOBProp := TraitementAffectationTrf(TForm(Ecran), TobAffectationTrf, TobMethode, TobListeEtab, 'M') ;

if (TOBProp=nil) or (TOBProp.detail.count=0) then
   begin
   if TOBProp<>nil then PgiInfo (TexteMessage[4], Ecran.Caption);
   end else
   begin
      // Consultation et/ou modification de la proposition d'affectation du transfert
      TheTOB:=TOBProp;
      CodeDepotEmet := GetControlText('CODDEPOT');
      AGLLanceFiche('MBO','PR_AFFTRANSFERT','','','ACTION=MODIFICATION;CODEDEPOTEMET='+CodeDepotEmet);
      TOBProp:=TheTOB;
      TheTOB:=nil;
   end;

// Enregistrement de la proposition d'affectation du transfert dans la Table
if (TOBProp<>nil) and (TOBProp.Detail.Count > 0) then
   begin
   TOBProp.SetAllModifie(True);
   TOBProp.InsertOrUpdateDB (True);
   CodePropTrf := GetControlText('CODPROPAFF');
   ExecuteSQL('Update PROPTRANSFENT Set GTE_ENCOURSPTRF="X" where GTE_CODEPTRF="'+CodePropTrf+'"') ;
   end;

TOBProp.Free ;
TobListeEtab.free;
TobListeEtab:=Nil;
TobAffectationTrf.Free ;
TobAffectationTrf:=Nil ;
end ;

procedure TOF_MBOTRANSFREF.OnClose ;
begin
  Inherited ;
  TobMethode.Free ;
  TobMethode:=Nil ;
  TheTob:=Nil ;
end ;

procedure TOF_MBOTRANSFREF.OnChangeCodeArtProp;
var QQ : TQuery ;
begin
  // Recherche de la désignation de l'article sélectionné
  QQ:=OpenSQL('Select GA_LIBELLE from ARTICLE where GA_CODEARTICLE="'+
              GetControlText('CODEARTICLE')+'" and GA_STATUTART<>"DIM"',True);
  if Not QQ.EOF then
       SetControlText('DESIGNATION', QQ.Findfield('GA_LIBELLE').AsString)
  else SetControlText('DESIGNATION', '');
  Ferme(QQ);
end ;

procedure TOF_MBOTRANSFREF.OnClickChoixMethode;
begin
TobMethode.ClearDetail ;
TheTob := TobMethode;
AglLanceFiche ('MBO','PR_METHAFFECT','','',GetControltext('CODDEPOT')) ;
TobMethode := TheTob;
if TobMethode.detail.count>0 then
   begin
   THValComboBox(GetControl('METHODE')).ReLoad ;
   SetControlText('METHODE', TobMethode.Detail[0].GetValue('CODE')) ;
   end;
end ;

Procedure AGLOnChangeCodeArtProp (Parms: array of variant; nb: integer) ;
var F : TForm;
    TOTOF : TOF ;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFVierge)
    then TOTOF := TFVierge(F).LaTof
    else exit;
  if (TOTOF is TOF_MBOTRANSFREF) then TOF_MBOTRANSFREF(TOTOF).OnChangeCodeArtProp else exit;
end;

Procedure AGLOnClickChoixMethode (Parms: array of variant; nb: integer) ;
var F : TForm;
    TOTOF : TOF ;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFVierge)
    then TOTOF := TFVierge(F).LaTof
    else exit;
  if (TOTOF is TOF_MBOTRANSFREF) then TOF_MBOTRANSFREF(TOTOF).OnClickChoixMethode else exit;
end;

Initialization
  registerclasses ( [ TOF_MBOTRANSFREF ] ) ;
  RegisterAglProc('OnChangeCodeArtProp', True , 1, AGLOnChangeCodeArtProp) ;
  RegisterAglProc('OnClickChoixMethode', True , 1, AGLOnClickChoixMethode) ;
end.
