{***********UNITE*************************************************
Auteur  ...... : Verdier
Créé le ...... : 21/06/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPVENTILSOUPLAN ()
Mots clefs ... : TOF;CPVENTILSOUPLAN
*****************************************************************}
Unit CPVENTILSOUPLAN_TOF ;

Interface

Uses StdCtrls, Controls, Classes, db, forms, sysutils, dbTables, ComCtrls,
     HCtrls, HEnt1, HMsgBox, UTOF, HTB97, HQry, UTob, HDB, Mul, Paramsoc, General,
     Ent1 ;

Type
  TOF_CPVENTILSOUPLAN = Class (TOF)
    Axe : ThValComboBox ;
    Da1 : THedit ;
    Da2 : THedit  ;
    Journal1 : ThValComboBox ;
    Journal2 : ThValComboBox ;
    Numeropiece1 : THedit ;
    Numeropiece2 : THedit ;
    Odana : TCheckBox ;
    Exo : ThValComboBox ;
    TobAna : Tob ;
    TobSouSec : Tob ;
    BOuvrir : TToolBarButton97 ;
    BCherche : TToolBarButton97 ;
    Fiche : TfMul ;
    ListeAna : THDBGrid ;
    QListeAna : TQuery ;
    Pages : TpageControl ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    procedure ExoOnChange(Sender: TObject);
    procedure OdanaOnClick(Sender: TObject);
    procedure BOuvrirOnClick(sender: TObject);
    procedure MarqueOrigine;
    procedure PrechargeOrigines (Jal:String;date1,Date2:TDatetime;DerniereDate:Boolean) ;
    procedure ChargeTobAna (Jal:String;date1,Date2:TDatetime;DerniereDate:Boolean) ;
    procedure VireInutiles;
    procedure TraitementLigne(TA: Tob);
    function ExisteSousSection(TSA: TSousPlan; TA: Tob ;Section : String): Boolean;
    procedure CreeSousSection(Axe, St:String ; TSA : TSousPlan)  ;
    Function LanceLeTraitement(Jal:String ; Date1,Date2:TDateTime ; DerniereDate:Boolean) : Boolean;
    function SwapGeneNew: Boolean;
  end ;

Implementation

var Titre : String ;

procedure TOF_CPVENTILSOUPLAN.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPVENTILSOUPLAN.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPVENTILSOUPLAN.BOuvrirOnClick (sender : TObject) ;
var i : Integer ;
    TobA :Tob ;
begin
ListeAna := Fiche.FListe ;
QListeAna := Fiche.Q ;

If PGIASK(TraduireMemoire('Confirmez-vous le traitement ?'),TraduireMemoire(Titre)) = MrNo then exit ;
If (ListeAna.NbSelected = 0) and (not ListeAna.AllSelected) then
  begin
  PGIInfo(TraduireMemoire('Vous n''avez sélectionné aucune ligne. Traitement impossible.'), TraduireMemoire(Titre)) ;
  Exit ;
  end ;

EnableControls(Fiche, False) ;

if SwapGeneNew then PgiInfo(TraduireMemoire('Traitement termine'),TraduireMemoire(Titre)) 
               else PgiBox(TraduireMemoire('Traitement interrompu. Relancer le traitement'),TraduireMemoire(Titre)) ;

BCherche.Click ;
EnableControls(Fiche, True) ;
end ;


Function TOF_CPVENTILSOUPLAN.LanceLeTraitement(Jal:String ; Date1,Date2:TDateTime ; DerniereDate:Boolean) : Boolean ;
var i : Integer ;
    TobA :Tob ;
begin
Result:=True ;
Try
  BeginTrans ;
    TobAna.ClearDetail ;
    ChargeTobAna (Jal,date1,Date2,DerniereDate);
    TobSouSec.ClearDetail ;
    for i:=0 to TobAna.Detail.Count-1 do
      begin
      TobA:=TobAna.detail[i] ;
      TraitementLigne (TobA) ;
      end ;
    TobAna.UpdateDb ; TobSouSec.InsertDb(nil) ;
    CommitTrans ;
Except
  RollBack ;
  Result:=False ;
end ;
End ;


Function TOF_CPVENTILSOUPLAN.SwapGeneNew : Boolean ;
Var DMin,DMax,DD1,DD2  : TDateTime ;
    ListeJ : TStrings ;
    Q,QExo : TQuery ;
    i,iper,Delta : integer ;
    DernDate,OkOk : Boolean ;
BEGIN
// Lecture des journaux
ListeJ:=TStringList.Create ; Result:=TRUE ; OkOk:=TRUE ;
Q:=OpenSQL('Select J_JOURNAL from JOURNAL ',true)  ; 
While Not Q.EOF do BEGIN ListeJ.Add(Q.Fields[0].AsString) ; Q.Next ; END ;
Ferme(Q) ;
// Balayage des écritures avec découpe
//InitMove(100,'') ;
for i:=0 to ListeJ.Count-1 do
    BEGIN
    QExo:=OpenSQl('Select EX_EXERCICE, EX_DATEDEBUT,EX_DATEFIN from EXERCICE Where EX_EXERCICE="'+Exo.Value+'" ',True) ;
    While Not QExo.EOF do
       BEGIN
       DMin:=QExo.Fields[1].AsDateTime ; DMax:=QExo.Fields[2].AsDateTime ; Delta:=Round((DMax-DMin)/10) ;
       for iper:=1 to 10 do
           BEGIN
//           MoveCur(FALSE) ;
           DernDate:=FALSE ;
           if iper<10 then BEGIN DD1:=DMin+(iper-1)*Delta ; DD2:=DD1+Delta ; END
                      else BEGIN DD1:=DMin+(iper-1)*Delta ; DD2:=DMax ; DernDate:=TRUE ; END ;
           OkOk:=LanceLeTraitement(ListeJ[i],DD1,DD2,DernDate) ;
           If Not OkOk Then Result:=FALSE ;
           END ;
       QExo.Next ;
       If Not OkOk Then Break ;
       END ;
    Ferme(QExo) ;
    If Not OkOk Then Break ;
    END ;
ListeJ.Clear ; ListeJ.Free ; //FiniMove ;
END ;

{If Jal<>'' Then St:=St+' WHERE E_JOURNAL<>"'+W_W+'" ' ;
If Jal<>'' Then St:=St+' AND E_JOURNAL="'+Jal+'" ' ;
If DD1<>0 Then St:=St+' AND E_DATECOMPTABLE>="'+UsDAteTime(DD1)+'" ' ;
If DD2<>0 Then
  BEGIN
  If DernDate Then St:=St+' AND E_DATECOMPTABLE<="'+UsDAteTime(DD2)+'" '
              Else St:=St+' AND E_DATECOMPTABLE<"'+UsDAteTime(DD2)+'" ' ;
  END ;
Q:=OpenSQL(st,TRUE) ;}


{***********A.G.L.Privé.*****************************************
Auteur  ...... : verdier
Créé le ...... : 25/06/2002
Modifié le ... : 25/06/2002
Description .. : Traitement d'1 ligne de la tob analytique
Mots clefs ... :
*****************************************************************}
procedure TOF_CPVENTILSOUPLAN.TraitementLigne (TA : Tob );
var Section,Axe : String ;
    TSA : TSousPlan ;
    j : Integer ;
    LeFicBase : TFichierBase ;
begin

Section:=Ta.GetValue('Y_SECTION') ;
Axe:=TA.GetValue('Y_AXE') ; LeFicBase:=AxeToFb(Axe) ;

For j:=1 to 6 do
  begin
  TSA:=VH^.SousPlanAxe[leficbase,j] ;
  If TSA.Code='' then Exit ;
  TA.PutValue('Y_SOUSPLAN'+IntToStr(j), copy(Section,TSA.Debut,TSa.Longueur)) ;
  if (not ExisteSousSection(TSA,TA,Section)) then CreeSousSection (Axe, copy(Section,TSA.Debut,TSa.Longueur), TSA);
  end ;
end ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : verdier
Créé le ...... : 25/06/2002
Modifié le ... :   /  /
Description .. : Verification de l'existence de la sous section
Mots clefs ... :
*****************************************************************}
Function TOF_CPVENTILSOUPLAN.ExisteSousSection (TSA : TSousPlan ; TA : Tob ; Section : String) : Boolean ;
var i,j: Integer ;
    S1,S2: String ;
begin
Result:=False ;
for i:=0 to TSA.ListeSP.Count-1 do
  begin
  S1:=TSA.ListeSP.Strings[i] ;
  S1:=ReadTokenSt(S1) ;
  if S1 = copy(Section,TSA.Debut,TSA.Longueur) then begin result:=True ; exit ; end ;
  end ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : verdier
Créé le ...... : 25/06/2002
Modifié le ... :   /  /
Description .. : Création de la sous section dans la table SSStrucr et
Suite ........ : dans la valeur Halley TSousPlan
Mots clefs ... :
*****************************************************************}
procedure TOF_CPVENTILSOUPLAN.CreeSousSection(Axe,St:String ; TSA : TSousPlan)  ;
var TobSS :Tob ;
    Str:String ;
begin
//Dans la table
TobSS:=Tob.Create('SSSTrucr',TobSouSec,-1) ;
TobSS.AddChampSup('PS_AXE',False) ; TobSS.PutValue('PS_AXE',Axe) ;
TobSS.AddChampSup('PS_SOUSSECTION',False) ; TobSS.PutValue('PS_SOUSSECTION', TSA.Code) ;
TobSS.AddChampSup('PS_CODE',False) ; TobSS.PutValue('PS_CODE',St) ;
TobSS.AddChampSup('PS_LIBELLE',False) ; TobSS.PutValue('PS_LIBELLE',St+' A Completer') ;
//Dans la variable Halley
Str:=St+';'+St+' A Completer' ;
TSA.ListeSP.Add(Str) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : verdier
Créé le ...... : 25/06/2002
Modifié le ... :   /  /    
Description .. : Chargement de la Tob Analytique
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPVENTILSOUPLAN.PrechargeOrigines (Jal:String;date1,Date2:TDatetime;DerniereDate:Boolean);
Var Q : TQuery ;
    ReqSql, St : String ;
begin
  ReqSql := 'SELECT Y_JOURNAL,Y_EXERCICE,Y_DATECOMPTABLE,Y_NUMEROPIECE,Y_NUMLIGNE,Y_AXE,Y_NUMVENTIL,Y_QUALIFPIECE, Y_SECTION, ' ;
  ReqSql := ReqSql + 'Y_SOUSPLAN1,Y_SOUSPLAN2,Y_SOUSPLAN3,Y_SOUSPLAN4,Y_SOUSPLAN5,Y_SOUSPLAN6 FROM ANALYTIQ ' ;
  St:=AnsiUpperCase(RecupWhereCritere(Pages)) ;
  if st='' then st:=' where y_journal="'+jal+'" ' else st:=St+' and Y_JOURNAL="'+jal+'" ' ;
  If Date1<>0 Then St:=St+' AND Y_DATECOMPTABLE>="'+UsDAteTime(Date1)+'" ' ;
  If Date2<>0 Then
  BEGIN
  If DerniereDate Then St:=St+' AND Y_DATECOMPTABLE<="'+UsDAteTime(Date2)+'" '
                  Else St:=St+' AND Y_DATECOMPTABLE<"'+UsDAteTime(Date2)+'" ' ;
  END ;
  ReqSql := ReqSql + St + ' ORDER BY Y_AXE,Y_SECTION' ;
  Q := OpenSQL(ReqSql, True) ;
  TobAna.LoadDetailDB('ANALYTIQ','','',Q,False,False) ;
  Ferme(Q) ;
  if TobAna.Detail.Count>0 then TobAna.Detail[0].AddChampSup('MARQUE',True) ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : verdier
Créé le ...... : 25/06/2002
Modifié le ... :   /  /    
Description .. : Marque Les Lignes de la Tob analytiques selectionnées 
Suite ........ : dans la grille de la mul
Mots clefs ... :
*****************************************************************}
procedure TOF_CPVENTILSOUPLAN.MarqueOrigine ;
var TobC : Tob ;
begin
  TOBC := TobAna.FindFirst(['Y_JOURNAL','Y_EXERCICE','Y_DATECOMPTABLE','Y_NUMEROPIECE','Y_NUMLIGNE','Y_AXE','Y_NUMVENTIL','Y_QUALIFPIECE'],
                           [QListeAna.FindField('Y_JOURNAL').AsString,
                            QListeAna.FindField('Y_EXERCICE').AsString,
                            QListeAna.FindField('Y_DATECOMPTABLE').AsDateTime,
                            QListeAna.FindField('Y_NUMEROPIECE').AsInteger,
                            QListeAna.FindField('Y_NUMLIGNE').AsInteger,
                            QListeAna.FindField('Y_AXE').AsString,
                            QListeAna.FindField('Y_NUMVENTIL').AsInteger,
                            QListeAna.FindField('Y_QUALIFPIECE').AsString],False) ;
  While TobC <> nil do
    begin
    TOBC.PutValue('MARQUE','X') ;
    TOBC := TobAna.FindNext(['Y_JOURNAL','Y_EXERCICE','Y_DATECOMPTABLE','Y_NUMEROPIECE','Y_NUMLIGNE','Y_AXE','Y_NUMVENTIL','Y_QUALIFPIECE'],
                             [QListeAna.FindField('Y_JOURNAL').AsString,
                              QListeAna.FindField('Y_EXERCICE').AsString,
                              QListeAna.FindField('Y_DATECOMPTABLE').AsDateTime,
                              QListeAna.FindField('Y_NUMEROPIECE').AsInteger,
                              QListeAna.FindField('Y_NUMLIGNE').AsInteger,
                              QListeAna.FindField('Y_AXE').AsString,
                              QListeAna.FindField('Y_NUMVENTIL').AsInteger,
                              QListeAna.FindField('Y_QUALIFPIECE').AsString],False) ;
    end ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : verdier
Créé le ...... : 25/06/2002
Modifié le ... :   /  /
Description .. : Vire les lignes de la tob non comprises dans la grille de la 
Suite ........ : mul
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPVENTILSOUPLAN.VireInutiles ;
var idxCpt : integer ;
    TobA : Tob ;
begin
  for idxCpt := TobAna.Detail.Count -1 downto 0 do
    begin
    TobA := TobAna.Detail[idxCpt] ;
    if TobA.GetValue('MARQUE') <> 'X' then begin TobA.Free ; TobA := Nil ; end ;
    end ;
end ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : verdier
Créé le ...... : 25/06/2002
Modifié le ... : 25/06/2002
Description .. : Chargement Tob Analytique par rapport  à la grille de la mul
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPVENTILSOUPLAN.ChargeTobAna (Jal:String;date1,Date2:TDatetime;DerniereDate:Boolean) ;
var idxCpt : integer ;
begin
  TobAna.ClearDetail ;
  PrechargeOrigines (Jal,date1,Date2,DerniereDate); ;
  if Not ListeAna.AllSelected then
    begin
    for idxCpt:=0 to ListeAna.NbSelected-1 do
      begin
      ListeAna.GotoLeBookmark(idxCpt) ;
      MarqueOrigine ;
      end ;
    VireInutiles ;
    end ;
end ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : verdier
Créé le ...... : 25/06/2002
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPVENTILSOUPLAN.OnArgument (S : String ) ;
begin
Fiche := TFMul(Ecran) ;
ListeAna := Fiche.FListe ;
QListeAna := Fiche.Q ;

Titre:='Ventilation des sous plans' ;

Axe:=ThValComboBox(GetControl('Y_AXE')) ;
Da1:=THedit(GetControl('Y_DATECOMPTABLE')) ;
Da2:=THedit(GetControl('Y_DATECOMPTABLE_')) ;
Journal1:=ThValComboBox(GetControl('Y_JOURNAL')) ;
Journal2:=ThValComboBox(GetControl('Y_JOURNAL_')) ;
Numeropiece1:=THedit(GetControl('Y_NUMEROPIECE')) ;
Numeropiece2:=THedit(GetControl('Y_NUMEROPIECE_')) ;
Odana:=TCheckBox(GetControl('ODANA')) ;
Exo:=ThValComboBox(GetControl('Y_EXERCICE')) ;
BOuvrir:=TToolBarButton97(GetControl('BOUVRIR')) ;
BCherche:=TToolBarButton97(GetControl('BCHERCHE')) ;
Pages:=TPageControl(GetControl('PAGES'));


Exo.OnChange := ExoOnChange  ;
Exo.Value:=VH^.EnCours.Code ;
ExoOnChange (nil);
SetControlText('Y_DATECOMPTABLE',DateToStr(Date));
SetControlText('Y_DATECOMPTABLE_',DateToStr(Date));


Odana.OnClick:=OdanaOnClick ;
BOuvrir.OnClick:=BOuvrirOnClick ;

TobAna:=Tob.Create('_Analytiq',nil,-1) ;
TobSouSec:=Tob.Create('_SSStrucr',Nil,-1) ;

  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : verdier
Créé le ...... : 25/06/2002
Modifié le ... :   /  /    
Description .. : Selection des differents types d'écritures analytiques
Mots clefs ... : 
*****************************************************************}
Procedure TOF_CPVENTILSOUPLAN.OdanaOnClick (Sender  : TObject ) ;
begin
SetControlText('XX_WHERE1','') ;
If GetCheckBoxState('ODANA')= CbChecked then SetControlText('XX_WHERE1','Y_TYPEANALYTIQUE="X"')
  else if GetCheckBoxState('ODANA')= CbUnChecked then SetControlText('XX_WHERE1','Y_TYPEANALYTIQUE="-"') ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : verdier
Créé le ...... : 25/06/2002
Modifié le ... :   /  /
Description .. : Changement d'exercice
Mots clefs ... : 
*****************************************************************}
Procedure TOF_CPVENTILSOUPLAN.ExoOnChange (Sender  : TObject ) ;
begin
if Exo.Value=VH^.EnCours.Code then
  begin
  Da1.Text:=DateToStr(VH^.EnCours.Deb) ;
  Da2.Text:=DateToStr(VH^.EnCours.Fin) ;
  end else
  begin
  Da1.Text:=DateToStr(VH^.Suivant.Deb) ;
  Da2.Text:=DateToStr(VH^.Suivant.Fin) ;
  end ;
end ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : verdier
Créé le ...... : 25/06/2002
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPVENTILSOUPLAN.OnClose ;
begin
If TobAna <> nil then begin TobAna.Free ; TobAna:=Nil ; end ;
If TobSouSec <> nil then begin TobSouSec.Free ; TobSouSec:=Nil ; end ;
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_CPVENTILSOUPLAN ] ) ;
end.
