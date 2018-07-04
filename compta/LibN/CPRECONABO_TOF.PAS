{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 16/03/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPRECONABO ()
Mots clefs ... : TOF;CPRECONABO
*****************************************************************}
Unit CPRECONABO_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul,
     FE_Main,
     HDB,
{$else}
     eMul, 
     uTob,
     MaineAGL,
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
  //   Echeance,
     Ent1,
     AGLInit,
     CONTABON_TOM,
     RapSuppr,
     HTB97,
     ED_TOOLS, // VideListe
     UTOF ; 

Type
  TOF_CPRECONABO = Class (TOF)
  private
    FNbLig   : Integer ;
    FTDelAbo : TList ;
    {$IFDEF EAGLCLIENT}
    FListe   : THGrid ;
    {$ELSE}
    FListe   : THDBGrid ;
    {$ENDIF}
   procedure FListeDblClick(Sender: TObject) ;
   procedure BOuvrirClick(Sender: TObject) ;
   procedure BSupprimeClick(Sender: TObject);
  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

function  ModifAboGene ( Code : String ; DateGene : TDateTime ) : boolean ;
procedure ReconductionAbonnements ;

Implementation

uses
 ULibEcriture ,
 UtilPgi ;

function  ModifAboGene ( Code : String ; DateGene : TDateTime ) : boolean ;
Var
 DD      : TDateTime ;
 LD      : TDateTime ;
 DModif  : TDateTime ;
 Q       :  TQuery ;
 Recond  : String ;
 Gen     : Integer ;
 Rep     : Integer ;
 Sep     : String3 ;
 Arr     : String3 ;
begin

 Result := False ;
 Q      := OpenSQL('Select * from CONTABON where CB_COMPTABLE="X" AND CB_CONTRAT="'+Code+'"',TRUE) ;

 if Q.EOF then
   begin
   Ferme(Q) ;
   Exit ;
   end ;

 Recond := Q.FindField('CB_RECONDUCTION').AsString ;
 DD     := Q.FindField('CB_DATEDERNGENERE').AsDateTime ;
 DModif := Q.FindField('CB_DATEMODIF').AsDateTime ;
 Rep    := Q.FindField('CB_NBREPETITION').AsInteger ;
 Gen    := Q.FindField('CB_DEJAGENERE').AsInteger ;
 Sep    := Q.FindField('CB_SEPAREPAR').AsString ;
 Arr    := Q.FindField('CB_ARRONDI').AsString ;
 Ferme(Q) ;

 if Recond ='SUP' then Exit ;
 if ((Rep=Gen) and (Recond<>'TAC')) then Exit ;

 LD := DD ;
 DD := ProchaineDate(DD,SEP,ARR) ;
 if DD>=DateGene then Exit ;

 While DD<DateGene do
   begin
   if (Recond='TAC') then
     begin
     if Rep=Gen then
       begin
       Gen:=1 ;
       DModif:=LD ;
       end
     else Inc(Gen) ;
     end ;
   LD := DD ;
   DD := ProchaineDate(DD,SEP,ARR) ;
   end ;

  ExecuteSQL('UPDATE CONTABON SET CB_DEJAGENERE="'     + IntToStr(Gen)      + '", '
                               + 'CB_DATEDERNGENERE="' + UsDateTime(LD)     + '", '
                               + 'CB_DATEMODIF="'      + UsDateTime(DModif) + '"'
                        + ' WHERE CB_COMPTABLE="X" AND CB_CONTRAT="'+Code+'"') ;
  Result := True ;

end;

procedure ReconductionAbonnements ;
begin

 if _Blocage(['nrCloture','nrBatch'],False,'nrAucun') then Exit ;

 AGLLanceFiche('CP','CPRECONABO','','','') ;

end ;

procedure TOF_CPRECONABO.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPRECONABO.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPRECONABO.OnUpdate ;
begin
 Inherited ;
end ;

procedure TOF_CPRECONABO.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPRECONABO.OnArgument (S : String ) ;
begin

 Inherited ;

 SetControlText('CB_DATECREATION'    , StDate1900 ) ;
 SetControlText('CB_DATECREATION_'   , StDate2099 ) ;
 SetControlText('CB_DATECONTRAT'     , StDate1900 ) ;
 SetControlText('CB_DATECONTRAT_'    , StDate2099 ) ;
 SetControlText('CB_DATEMODIF'       , StDate1900 ) ;
 SetControlText('CB_DATEMODIF_'      , StDate2099 ) ;
 SetControlText('CB_DATEDERNGENERE'  , StDate1900 ) ;
 SetControlText('CB_DATEDERNGENERE_' , StDate2099 ) ;

 TToolbarButton97(GetControl('BSUPP')).OnClick := BSupprimeClick ;
 TToolbarButton97(GetControl('BOUVRIR')).OnClick := BOuvrirClick ;
// TToolbarButton97(GetControl('bDelete')).OnClick := OnSupprimeClick;


 {$IFDEF EAGLCLIENT}
  FListe  := THGrid (GetControl ('FLISTE')) ;
 {$ELSE}
  FListe := THDBGrid (GetControl ('FLISTE'));
 {$ENDIF}

  FTDelAbo := TList.Create ;

  FListe.OnDblClick := FListeDblClick ;

end ;

procedure TOF_CPRECONABO.OnClose ;
begin

 Videliste(FTDelAbo) ;
 FTDelAbo.Free ;

 Inherited ;
 
end ;

procedure TOF_CPRECONABO.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPRECONABO.OnCancel () ;
begin
  Inherited ;
end ;


procedure TOF_CPRECONABO.FListeDblClick(Sender: TObject);
begin
  {$IFDEF EAGLCLIENT}
  TheMulQ := TFMul(Ecran).Q.TQ ;
 {$ELSE}
  TheMulQ := TFMul(Ecran).Q ;
 {$ENDIF}
 if TheMulQ.EOF and TheMulQ.Bof then Exit ;
 ParamAbonnement(True,GetField('CB_CONTRAT'),taModif) ;
end;

procedure TOF_CPRECONABO.BOuvrirClick(Sender: TObject);
var i        : Integer ;
    lStCode  : String ;
    lDtDate  : TDateTime;
begin

 FNbLig := Fliste.NbSelected ;
 if FNbLig<=0 then BEGIN  PGIBox('Vous n''avez sélectionné aucun abonnement',Ecran.Caption) ; Exit ; END ;
  if PGIAsk('Confirmez-vous la reconduction des abonnements sélectionnés ?','')<>mrYes then Exit ;

 try

     BeginTrans ;

     for i:=0 to FNbLig-1 do
       begin
       Fliste.GotoLeBookMark(i) ;

       lStCode := GetField('CB_CONTRAT') ;        // Code du contrat         
       lDtDate := GetField('CB_DATEDERNGENERE') ; // Date du contrat

       ExecuteSQL('UPDATE CONTABON SET CB_DEJAGENERE=0, CB_DATEMODIF="' + UsDateTime( Date ) + '",'
                                   + ' CB_DATECONTRAT="' + UsDateTime( lDtDate ) +
                   '" WHERE CB_COMPTABLE="X" AND CB_CONTRAT="' + lStCode + '" '
                     + 'AND CB_RECONDUCTION<>"SUP" ') ;
       end ;

     CommitTrans ;
   Except
   rollback ;
   end;

 TFMul(ECRAN).BChercheClick( nil );
 PGIBox('La reconduction des abonnements a été effectuée avec succès',Ecran.Caption) ;

end;

procedure TOF_CPRECONABO.BSupprimeClick(Sender: TObject);
Var i    : Integer ;
    X    : DelInfo ;
    Code : String ;
    Lib  : String ;
begin

 FNbLig := Fliste.NbSelected ;
 FTDelAbo.Clear ;

 if FNbLig<=0 then
   begin
   PGIBox('Vous n''avez sélectionné aucun abonnement',Ecran.Caption) ;
   Exit ;
   end ;
 if PGIAsk('Confirmez-vous la suppression des abonnements sélectionnés ?','')<>mrYes then Exit ;

 try

   BeginTrans ;

   for i:=0 to FNbLig-1 do
     begin
     Fliste.GotoLeBookMark(i) ;
     Code := GetField('CB_CONTRAT') ;
     Lib  := GetField('CB_LIBELLE') ;
     X    := DelInfo.Create ;
     X.LeCod   := Code ;
     X.LeLib   := Lib ;
     X.LeMess  := TraduireMemoire('Révision des abonnements') ;
     FTDelAbo.Add(X) ;
     ExecuteSQL('DELETE FROM CONTABON WHERE CB_COMPTABLE="X" AND CB_CONTRAT="' + Code + '"') ;
     end ;

   CommitTrans ;

   except
     Rollback ;
   end ;

 TFMul(ECRAN).BChercheClick( nil );
 if PGIAsk('Voulez-vous voir la liste de abonnements détruits ?','')=mrYes then RapportDeSuppression(FTDelAbo,1) ;

end;



Initialization
  registerclasses ( [ TOF_CPRECONABO ] ) ; 
end.
