{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 20/12/2001
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : SECTION (SECTION)
Mots clefs ... : TOM;SECTION
*****************************************************************}
Unit SECTION_TOM ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     dbtables,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOM,
     Fiche,
     FichList,
     UTob
     ,Ent1
     , HDB
     ,HTB97
     ,Menus
     ,CritEdt
     , FE_Main
     ,UtilFichCpta
     , HCompte
     ;


Type
  TOM_SECTION = Class (TOM)
    SoldeP, SoldeE, SoldeS : THNumEdit ;
    PMenu : TPopupMenu ;
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
  private
    isChantier : Boolean ;
    Titre : string ;
    LeFb             : String ;
    QuelAxe          : String3 ;
    LgCode : Byte ;
    procedure FormateLesMontants;
    procedure EstUnAxeChantier(QuelAxe : string);
    function VerifCoherenceTL: Boolean;
    procedure BBalSecGenOnClick(Sender: TObject);
    procedure BZEcriMvtOnClick(Sender: TObject);
    procedure BCumulOnClick(Sender: TObject);
    procedure SwapttRepart;
    procedure Ergo35;
    procedure BCodSecOnClick(Sender: TObject);
    procedure BSaiCodsecClick(Sender: TObject);
    procedure SectionOnEnter(Sender: TObject);
    procedure SectionOnExit(Sender: TObject);
    function OkConfidentiel: Boolean;
    end ;

Implementation

uses
     {umMens,}ZoomAna, UTILEDT, QRBLSEGE,Structur ;

Procedure TOM_SECTION.FormateLesMontants ;
begin
  SetControlProperty('S_DEBITDERNMVT','DisPlayFormat',StrfMask(V_PGI.OkdecV,'',True)) ;
  SetControlProperty('S_CREDITDERNMVT','DisPlayFormat',StrfMask(V_PGI.OkdecV,'',True)) ;
  SetControlProperty('S_TOTALDEBIT','DisPlayFormat',StrfMask(V_PGI.OkdecV,'',True)) ;
  SetControlProperty('S_TOTALCREDIT','DisPlayFormat',StrfMask(V_PGI.OkdecV,'',True)) ;
  SetControlProperty('S_TOTDEBP','DisPlayFormat',StrfMask(V_PGI.OkdecV,'',True)) ;
  SetControlProperty('S_TOTCREP','DisPlayFormat',StrfMask(V_PGI.OkdecV,'',True)) ;
  SetControlProperty('S_TOTDEBE','DisPlayFormat',StrfMask(V_PGI.OkdecV,'',True)) ;
  SetControlProperty('S_TOTCREE','DisPlayFormat',StrfMask(V_PGI.OkdecV,'',True)) ;
  SetControlProperty('S_TOTDEBS','DisPlayFormat',StrfMask(V_PGI.OkdecV,'',True)) ;
  SetControlProperty('S_TOTCRES','DisPlayFormat',StrfMask(V_PGI.OkdecV,'',True)) ;
  SetControlProperty('S_TOTDEBANO','DisPlayFormat',StrfMask(V_PGI.OkdecV,'',True)) ;
  SetControlProperty('S_TOTCREANO','DisPlayFormat',StrfMask(V_PGI.OkdecV,'',True)) ;
  ChangeMask(SoldeP,V_PGI.OkdecV,'') ;
  ChangeMask(SoldeE,V_PGI.OkdecV,'') ;
  ChangeMask(SoldeS,V_PGI.OkdecV,'') ;
end ;

Procedure TOM_SECTION.EstUnAxeChantier(QuelAxe: string) ;
begin
  IsChantier:=VH^.Cpta[AxeToFb(QuelAxe)].Chantier ;
  SetControlEnabled('TS_CHANTIER',IsChantier ); SetControlEnabled('S_CHANTIER',IsChantier) ;
  SetControlEnabled('TS_MAITREOEUVRE',IsChantier) ; SetControlEnabled('S_MAITREOEUVRE',IsChantier) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 09/01/2002
Modifié le ... :   /  /    
Description .. : Vérification des tables libres section.
Mots clefs ... : SECTION;LIBRE
*****************************************************************}
Function TOM_SECTION.VerifCoherenceTL : Boolean ;
Var i : Integer ;
    Alerte : Boolean ;
    tz : TZoomTable ;
begin
  Result:=True ;
  if Not TTabSheet(GetControl('ZL')).TabVisible then begin Result:=True ; Exit ; end ;
  Alerte:=False ;
  for i:=0 to 9 do
    begin
    Case i of
     0: tz := tzNatSect0 ;
     1: tz := tzNatSect1 ;
     2: tz := tzNatSect2 ;
     3: tz := tzNatSect3 ;
     4: tz := tzNatSect4 ;
     5: tz := tzNatSect5 ;
     6: tz := tzNatSect6 ;
     7: tz := tzNatSect7 ;
     8: tz := tzNatSect8 ;
     9: tz := tzNatSect9 ;
    end ;
    if Not(THDBEdit(GetControl('S_TABLE'+IntToStr(i))).Enabled) then Continue ;
    if THDBEdit(GetControl('S_TABLE'+IntToStr(i))).Text='' then Continue ;
    if Not ExisteSQL('SELECT NT_NATURE FROM NATCPTE WHERE NT_TYPECPTE="'+tzToNature(tz)+'" AND NT_NATURE="'+THDBEdit(GetControl('S_TABLE'+IntToStr(i))).Text+'"') then Alerte := True;
    end ;
  if Alerte then if PGIAsk(TraduireMemoire('Certains codes de tables libres n''existent pas. Désirez-vous continuer ?'),'')<>mrYes then result := false ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 09/01/2002
Modifié le ... :   /  /    
Description .. : Paramètrage des tablettes suivant l'axe sélectionné.
Mots clefs ... :
*****************************************************************}
Procedure TOM_SECTION.SwapttRepart ;
begin
Case QuelAxe[2] of
  '1' : begin
        SetControlProperty('S_SECTION','DATATYPE', 'TZSECTION' ) ;
        SetControlProperty('S_CLEREPARTITION','DATATYPE', 'TTCLEREPART1' ) ;
        SetControlProperty('S_CORRESP1','DATATYPE', 'TZCORRESPSEC11' ) ;
        SetControlProperty('S_CORRESP2','DATATYPE', 'TZCORRESPSEC12' ) ;
        LeFb:='fbAxe1' ;
        end ;
  '2' : begin
        SetControlProperty('S_SECTION','DATATYPE', 'TZSECTION2' ) ;
        SetControlProperty('S_CLEREPARTITION','DATATYPE', 'TTCLEREPART2' ) ;
        SetControlProperty('S_CORRESP1','DATATYPE', 'TZCORRESPSEC21' ) ;
        SetControlProperty('S_CORRESP2','DATATYPE', 'TZCORRESPSEC22' ) ;
        LeFb:='fbAxe2' ;
        end ;
  '3' : begin
        SetControlProperty('S_SECTION','DATATYPE', 'TZSECTION3' ) ;
        SetControlProperty('S_CLEREPARTITION','DATATYPE', 'TTCLEREPART3' ) ;
        SetControlProperty('S_CORRESP1','DATATYPE', 'TZCORRESPSEC31' ) ;
        SetControlProperty('S_CORRESP2','DATATYPE', 'TZCORRESPSEC32' ) ;
        LeFb:='fbAxe3' ;
        end ;
  '4' : begin
        SetControlProperty('S_SECTION','DATATYPE', 'TZSECTION4' ) ;
        SetControlProperty('S_CLEREPARTITION','DATATYPE', 'TTCLEREPART4' ) ;
        SetControlProperty('S_CORRESP1','DATATYPE', 'TZCORRESPSEC41' ) ;
        SetControlProperty('S_CORRESP2','DATATYPE', 'TZCORRESPSEC42' ) ;
        LeFb:='fbAxe4' ;
        end ;
  '5' : begin
        SetControlProperty('S_SECTION','DATATYPE', 'TZSECTION5' ) ;
        SetControlProperty('S_CLEREPARTITION','DATATYPE', 'TTCLEREPART5' ) ;
        SetControlProperty('S_CORRESP1','DATATYPE', 'TZCORRESPSEC51' ) ;
        SetControlProperty('S_CORRESP2','DATATYPE', 'TZCORRESPSEC52' ) ;
        LeFb:='fbAxe5' ;
        end ;
  end ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 09/01/2002
Modifié le ... :   /  /    
Description .. : Edition de la Balance section / général
Mots clefs ... : 
*****************************************************************}
procedure TOM_SECTION.BBalSecGenOnClick ( Sender : TObject ) ;
Var Crit : TCritEdt ;
    D1,D2 : TdateTime ;
    Etab : String ;
begin
Fillchar(Crit,SizeOf(Crit),#0) ;
D1:=VH^.Encours.Deb ; D2:=VH^.Encours.Fin ;
If VH^.Entree.Code=VH^.Suivant.Code Then BEGIN D1:=VH^.Suivant.Deb ; D2:=VH^.Suivant.Fin ; END ;
Crit.Date1:=D1 ; Crit.Date2:=D2 ;
Crit.DateDeb:=Crit.Date1 ; Crit.DateFin:=Crit.Date2 ;
Crit.NatureEtat:=neBal ;
Crit.Bal.Axe:=GetControlText('S_Axe') ;
InitCritEdt(Crit) ;
Crit.Cpt1:=GetControlText('S_Section') ;
Crit.Cpt2:=Crit.Cpt1 ;
Etab:=EtabForce ; if Etab<>'' then Crit.Etab:=Etab ;
BLSEGEZoom(Crit) ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 09/01/2002
Modifié le ... :   /  /    
Description .. : Détail des mouvements de la section.
Mots clefs ... : 
*****************************************************************}
procedure TOM_SECTION.BZEcriMvtOnClick ( Sender : TObject ) ;
begin
{$IFNDEF IMP}
{$IFNDEF CCMP}
{$IFNDEF GCGC}
  AGLLanceFiche('CP','ZOOMANAL','','',';;'+GetControlText('S_SECTION')+';'+GetControlText('S_AXE')+';'+EXRF(VH^.Entree.Code)) ;
{$ENDIF}
{$ENDIF}
{$ENDIF}
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 09/01/2002
Modifié le ... :   /  /    
Description .. : Cumuls mensuels de la section.
Mots clefs ... :
*****************************************************************}
procedure TOM_SECTION.BCumulOnClick ( Sender : TObject ) ;
begin
//CumulCpteMensuel(LeFb,GetControlText('S_SECTION'),GetControlText('S_LIBELLE'),VH^.Entree) ;
  AGLLanceFiche('CP','CPCUMMENS','','',LeFb+';'+GetControlText('S_SECTION')+';'+GetControlText('S_LIBELLE')+';'+VH^.Entree.Code) ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 09/01/2002
Modifié le ... :   /  /    
Description .. : Gestion affichage des zones suivant la série ( S3, S5 ).
Mots clefs ... : 
*****************************************************************}
procedure TOM_SECTION.Ergo35 ;
begin
if ((EstSerie(S3)) or (EstSerie(S5))) then
   begin
   SetControlVisible('S_CORRESP2', false ) ;
   SetControlVisible('TS_CORRESP2', false ) ;
   SetControlVisible('S_CLEREPARTITION', false ) ;
   SetControlVisible('TS_CLEREPARTITION', false ) ;
   SetControlVisible('BEVCLEREPART', false ) ;
   if EstSerie(S3) then
      begin
      SetControlVisible('S_AXE', false ) ;
      SetControlVisible('TS_AXE', false ) ;
      SetControlVisible('S_CORRESP1', false ) ;
      SetControlVisible('TS_CORRESP1', false ) ;
      SetControlProperty('PCOMPLEMENT','TabVisible', false ) ;
      SetControlVisible('GBCONFIDENTIEL', false ) ;
      end ;
   end ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 09/01/2002
Modifié le ... :   /  /    
Description .. : Affichage de l'assistant de création de code section.
Mots clefs ... : 
*****************************************************************}
procedure TOM_SECTION.BCodSecOnClick (Sender : TObject) ;
var ArgRetour : string ;
begin
  ArgRetour := AGLLanceFiche('CP','CPASSSECSTRUC','','',QuelAxe) ;
  SetField('S_SECTION',ReadTokenPipe(ArgRetour,'|' )) ;
  If Length(ArgRetour) > 35 then ArgRetour := Copy(ArgRetour,1,35 ) ;
  SetField('S_LIBELLE',ArgRetour) ;
  SetFocusControl('S_SECTION') ;
end ;

procedure TOM_SECTION.BSaiCodsecClick(Sender: TObject);
begin
  if VH^.Cpta[AxeToFb(GetControlText('S_AXE'))].Structure then ParamPlanAnal(GetControlText('S_AXE')) ;
  SetFocusControl('S_SECTION') ;
//S_SECTION.SelStart:=Length(S_SECTION.Text) ; S_SECTION.SelLength:=0 ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 09/01/2002
Modifié le ... :   /  /
Description .. : Activation des boutons "Assistant" et "paramètrage
Suite ........ : structure" en entrant sur la zone Section.
Mots clefs ... :
*****************************************************************}
procedure TOM_SECTION.SectionOnEnter ( Sender : TObject ) ;
begin
  if DS.State = dsInsert then
    begin
    SetControlEnabled('BCODSEC',DS.state in [dsInsert]) ;
    SetControlEnabled('BSAICODSEC', GetControlEnabled('BCODSEC')) ;
    end ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 09/01/2002
Modifié le ... :   /  /
Description .. : Gestion affichage des zones en sortie de saisie du code, et
Suite ........ : affinage du code section saisi.
Mots clefs ... :
*****************************************************************}
procedure TOM_SECTION.SectionOnExit(Sender: TObject);
begin
  SetControlEnabled('BCODSEC',false) ;
  SetControlEnabled('BSAICODSEC', false) ;
  if GetControlText('S_SECTION')='' then BEGIN SetControlEnabled('S_AXE',True) ; Exit ; END ;
  SetControlEnabled('S_AXE',False) ;
  if Length(GetControlText('S_SECTION'))>LgCode then SetField('S_SECTION',Copy(GetControlText('S_SECTION'),1,LgCode)) ;
  if Length(GetControlText('S_SECTION'))<LgCode then SetField('S_SECTION', BourreLaDonc(GetControlText('S_SECTION'),AxeToFb(QuelAxe))) ;
end;

Function TOM_SECTION.OkConfidentiel : Boolean ;
BEGIN
  Result:=False ;
  if (V_PGI.Confidentiel='0') and (GetControlText('S_Confidentiel')='X') then
    BEGIN
    PGIBox(TraduireMemoire('L''accès à cette section est interdit.'),Titre) ;
    Exit ;
    END ;
  SetControlVisible('GBCONFIDENTIEL', (V_PGI.Confidentiel='1')) ;
  SetControlEnabled('GBCONFIDENTIEL', V_PGI.Superviseur) ;
  Result:=True ;
END ;

procedure TOM_SECTION.OnNewRecord ;
begin
  Inherited ;
  SetControlEnabled('BCODSEC',true ) ;
  SetControlEnabled('BSAICODSEC',true ) ;
  SetControlEnabled('BMENU',false ) ;

  THDBValComboBox(GetControl('S_SENS')).ItemIndex := 2 ; SetField('S_SENS', GetControlText('S_SENS')) ;
  THDBValComboBox(GetControl('S_AXE')).ItemIndex := 0 ; SetField('S_AXE', GetControlText('S_AXE')) ;

  SetField('S_SOLDEPROGRESSIF', 'X' ) ;
  SetField('S_CONFIDENTIEL', 0 ) ;

  SetFocusControl('S_SECTION') ;
  LgCode:=VH^.Cpta[AxeToFb(GetControlText('S_AXE'))].Lg ;

end ;

procedure TOM_SECTION.OnDeleteRecord ;
begin
  Inherited ;
{  if EstDansAnalytiq(GetControlText('S_SECTION'),GetControlText('S_AXE')) then
    begin
    LastErrorMsg := TraduireMemoire('Vous ne pouvez pas supprimer cette section : elle possède des écritures analytiques.') ;
    LastError := -1 ; Exit ;
    end ;
  if EstDansAxe(GetControlText('S_SECTION'),GetControlText('S_AXE')) then
    begin
    LastErrorMsg := TraduireMemoire('Vous ne pouvez pas supprimer cette section : elle est déclarée en tant que section d''attente pour cet axe.') ;
    LastError := -1 ; Exit ;
    end ;
  if EstCorresp(GetControlText('S_SECTION'),GetControlText('S_AXE')) then
    begin
    LastErrorMsg := TraduireMemoire('Vous ne pouvez pas supprimer cette section : elle est déclarée en tant que section de correspondance.') ;
    LastError := -1 ; Exit ;
    end ;
  if EstDansGuideAna(GetControlText('S_SECTION'),GetControlText('S_AXE')) then
    begin
    LastErrorMsg := TraduireMemoire('Vous ne pouvez pas supprimer cette section : elle est référencée dans un guide.') ;
    LastError := -1 ; Exit ;
    end ;
  if EstDansVentil(GetControlText('S_SECTION'),GetControlText('S_AXE')) then
    begin
    LastErrorMsg := TraduireMemoire('Vous ne pouvez pas supprimer cette section : elle est utilisée dans une ventilation type ou une ventilation par défaut d''un compte.') ;
    LastError := -1 ; Exit ;
    end ;
}
end ;

procedure TOM_SECTION.OnUpdateRecord ;
begin
  Inherited ;
  if GetControlText('S_SECTION')='' then
    begin
    LastErrorMsg := TraduireMemoire('Vous devez renseigner un code.') ;
    LastError := -1 ;
    SetFocusControl('S_SECTION') ;
    Exit ;
    end ;
  if not VerifCoherenceTL then begin LastError := -1 ; Exit ; end ;
  Case VerifCorrespondance(2+StrToInt(QuelAxe[2]),GetControlText('S_CORRESP1'),GetControlText('S_CORRESP2')) of
    0 : ;
    1 : begin LastErrorMsg := TraduireMemoire('La section de correspondance n°1 n''existe pas. Vous devez la modifier.') ; LastError := -1 ; Exit ; end ;
    2 : begin LastErrorMsg := TraduireMemoire('La section de correspondance n°2 n''existe pas. Vous devez la modifier.') ; LastError := -1 ; Exit ; end ;
    END ;
  if (DS.State = dsInsert) and (ExisteCpte(GetControlText('S_SECTION'),AxeToFb(QuelAxe))) then
   begin
   SetFocusControl('S_SECTION') ;
   LastError := -1 ;
   LastErrorMsg := TraduireMemoire('Le code que vous avez saisi existe déjà. Vous devez le modifier.') ;
   end ;

end ;

procedure TOM_SECTION.OnAfterUpdateRecord ;
begin
  Inherited ;
  if DS.State <> dsInsert then
    begin
    SetControlEnabled('S_SECTION', false ) ;
    SetControlEnabled('BMENU',true ) ;
    end ;
end ;

procedure TOM_SECTION.OnLoadRecord ;
begin
  Inherited ;
  if DS.State <> dsInsert then SetControlEnabled('S_SECTION',false ) ;

  if DS.State<>dsInsert then
    begin
    SetControlEnabled('BCODSEC',false ) ;
    SetControlEnabled('BSAICODSEC',false ) ;
    SetControlEnabled('BMENU',true ) ;
    end ;

  AfficheLeSolde(SoldeP,StrToFloat(GetField('S_TOTDEBP')),StrToFloat(GetField('S_TOTCREP'))) ;
  AfficheLeSolde(SoldeE,StrToFloat(GetField('S_TOTDEBE')),StrToFloat(GetField('S_TOTCREE'))) ;
  AfficheLeSolde(SoldeS,StrToFloat(GetField('S_TOTDEBS')),StrToFloat(GetField('S_TOTCRES'))) ;
  EstUnAxeChantier(GetField('S_AXE')) ;
  QuelAxe := GetControlText('S_AXE') ;
  LgCode:=VH^.Cpta[AxeToFb(QuelAxe)].Lg ;
  SwapttRepart ;
end ;

procedure TOM_SECTION.OnChangeField ( F: TField ) ;
begin
  Inherited ;
  if F.FieldName = 'S_AXE' then
    begin
    QuelAxe := GetControlText('S_AXE') ;
    LgCode:=VH^.Cpta[AxeToFb(QuelAxe)].Lg ;
    SwapttRepart ;
    end ;
end ;

procedure TOM_SECTION.OnArgument ( S: String ) ;
begin
   Inherited ;
  if Not OKConfidentiel then Exit ;
  Titre := Ecran.Caption ;
  SoldeP := THNumEdit(GetControl('SSOLDEP')) ;
  SoldeE := THNumEdit(GetControl('SSOLDEE')) ;
  SoldeS := THNumEdit(GetControl('SSOLDES')) ;
  LibellesTableLibre(TTabSheet(GetControl('ZL')),'TS_TABLE','S_TABLE','S') ;
  FormateLesMontants ;

  PMenu := TPopupMenu(GetControl('MENU')) ;
  PMenu.Items[0].OnClick := BBalSecGenOnClick ;
  PMenu.Items[1].OnClick := BZEcriMvtOnClick ;
  PMenu.Items[2].OnClick := BCumulOnClick ;

  TToolBarButton97(GetControl('BCODSEC')).onClick := BCodSecOnClick ;
  TToolBarButton97(GetControl('BSAICODSEC')).onClick := BSaiCodSecClick ;

  THDBEdit(GetControl('S_SECTION')).OnEnter := SectionOnEnter ;
  THDBEdit(GetControl('S_SECTION')).OnExit := SectionOnExit ;

  Ergo35 ;

end ;

procedure TOM_SECTION.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_SECTION.OnCancelRecord ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOM_SECTION ] ) ;
end.
