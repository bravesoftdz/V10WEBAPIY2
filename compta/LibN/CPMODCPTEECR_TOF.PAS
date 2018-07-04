{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 30/11/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : CPMODCPTEECR ()
Mots clefs ... : TOF;CPMODCPTEECR
*****************************************************************}
Unit CPMODCPTEECR_TOF ;

Interface

Uses StdCtrls, Controls, Classes, db, forms, sysutils, dbTables, ComCtrls,
     HCtrls, HEnt1, HMsgBox, UTOF
     , HQry, Ent1, HTB97, Mul, UTob, HDB,
     ParamSoc,		//GetParamSocSecur YMO
     Saisie
     ;

Const Titre = 'Modification des comptes dans les écritures' ;

Type
  TOF_CPMODCPTEECR = Class (TOF)
    Exercice : THValComboBox ;
    DateComptable, DateComptable_ : THEdit ;
    Fiche : TFMul ;
    ListeEcr : THDBGrid ;
    QListeEcr : TQuery ;
    Pages : TpageControl ;
    TobEcr, TobJal : Tob ;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    procedure ChoixCpteOnClick(sender: TObject);
    procedure EAuxiliaireOnChange(Sender: TObject);
    procedure EGeneralOnChange(Sender: TObject);
    procedure ExerciceOnChange(Sender: TObject);
    procedure BOuvrirOnClick(sender: TObject);
    procedure MarqueOrigine;
    procedure PrechargeOrigines (ModifLettrage : boolean) ;
    procedure RecupMultiSelect( ModifLettrage : boolean) ;
    procedure VireInutiles;
    function TraitementEcriture(ModifGene, ModifLettrage : boolean) : boolean;
    function TraiteCptBqe(Compte: string; var NewJal: string): boolean;
    procedure ChargeTobJal;
    procedure VireJal;
    procedure MajSolde(P, Signe: char; Code: string);
    function VerifPointageCpt(CptOri, CptDest: string): boolean;
    procedure ModifJal(TobE: Tob);
    procedure CalculCumul(Exo : string; Debit, Credit: double);
    procedure ToutEcrOnClick(Sender: TObject);
    procedure ModifLigneAnaAuxi(TobE: Tob; CptAuxi: String);
    procedure ChargeTobSection(var TobSection: Tob);
    procedure VireSection(TobSection: tob);
    procedure TraitementAnaGene(TobE, TobSection : Tob);
    procedure MajSection(TobS: Tob; Exo: String; TobA: Tob; Signe: integer);
    procedure RechContrepartie(var TobContre, TobAnaContre : tob; TobE: tob; NatCpte: string);
    procedure MajTobContre(var TobContre, TobAnaContre: tob; NatCpte,NewCpte: string);
    procedure MajJournal(CodeJal: string);
    procedure FListeOnDblClick(sender: TObject);
    procedure AuxiElipsisClick(Sender : TObject);
  end ;

Implementation

uses UTofMulParamGen; {13/04/07 YMO F5 sur Auxiliaire }

var NatureCompte : string ;
    CumDeb, CumCre, CumDebP, CumCreP, CumDebE, CumCreE, CumDebS, CumCreS : double ;
    VentilOri, VentilDes : boolean ;
    VentilCptOri, VentilCptDes : Array[fbAxe1..fbAxe5] of string ;

procedure TOF_CPMODCPTEECR.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPMODCPTEECR.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPMODCPTEECR.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPMODCPTEECR.OnLoad ;
begin
  // Affichage des écritures pour un compte donné obligatoire
  if (GetControlText('E_GENERAL')='') and (GetControlText('E_AUXILIAIRE')='') then
  begin
  SetControlText('XX_WHERE','1=2' ) ;
  exit ;
  end ;
  Case THRadioGroup(GetControl('CHOIXCPTE')).ItemIndex of
  0 :begin
     // pour un compte général, modif possible des écritures de l'exo en cours ou suivant hors a nouveau
     SetControlText('XX_WHERE','E_EXERCICE = "'+VH^.EnCours.Code+'"') ;
     if VH^.Suivant.Code <> '' then SetControlText('XX_WHERE','('+GetControlText('XX_WHERE')+' OR E_EXERCICE = "'+VH^.Suivant.Code+'")') ;
     SetControlText('XX_WHERE',GetControlText('XX_WHERE')+' AND E_ECRANOUVEAU="N"') ;
     end ;
     // pour compte auxiliaire, aucune contrainte sur les écritures
  1 :SetControlText('XX_WHERE','' ) ;
  end ;
  Inherited ;

end ;

procedure TOF_CPMODCPTEECR.ChoixCpteOnClick (sender : TObject ) ;
begin
  Case THRadioGroup(GetControl('CHOIXCPTE')).ItemIndex of
  0 :begin
     SetControlText('E_AUXILIAIRE','') ;
     SetControlEnabled('TE_AUXILIAIRE',false) ;
     SetControlEnabled('E_AUXILIAIRE',false) ;
     SetControlText('TOUTECR', '') ;
     SetControlEnabled('TOUTECR',false) ;
     EAuxiliaireOnChange(Self) ;

     SetControlEnabled('TE_GENERAL',true) ;
     SetControlEnabled('E_GENERAL',true) ;
     SetControlText('XX_WHERE','E_EXERCICE = "'+VH^.EnCours.Code+'"') ;
     if VH^.Suivant.Code <> '' then SetControlText('XX_WHERE',GetControlText('XX_WHERE')+' OR E_EXERCICE = "'+VH^.Suivant.Code+'"') ;
     end ;
  1 :begin
     SetControlText('E_GENERAL','') ;
     SetControlEnabled('TE_GENERAL',false) ;
     SetControlEnabled('E_GENERAL',false) ;
     EGeneralOnChange(Self) ;

     SetControlEnabled('TE_AUXILIAIRE',true) ;
     SetControlEnabled('E_AUXILIAIRE',true) ;
     SetControlEnabled('TOUTECR',true) ;
     SetControlText('XX_WHERE', '' ) ;
     end ;
   end ;
end ;

function RechNatureCpte (Typ : char; Compte : string) : string ;
var Q : TQuery ;
    Req, NomTable, ColCpt, ColNature : string ;
begin
  // recherche de la nature du compte saisi
  result := '' ;
  NomTable := 'GENERAUX' ; If Typ = 'T' then NomTable := 'TIERS' ;
  ColCpt := 'G_GENERAL'; If Typ = 'T' then ColCpt := 'T_AUXILIAIRE' ;
  ColNature := 'G_NATUREGENE' ; If Typ = 'T' then ColNature := 'T_NATUREAUXI' ;

  Req := 'SELECT '+ColNature+' FROM '+NomTable+' WHERE '+ColCpt+'="'+Compte+'"' ;
  Q := OpenSQL(Req, True) ;
  if not Q.Eof then result := Q.Fields[0].asString ;
  Ferme(Q) ;
end ;

procedure TOF_CPMODCPTEECR.EGeneralOnChange (Sender : TObject ) ;
Var NomTable : string ;
    AncNature : string ;
begin
  AncNature := NatureCompte ;
  if GetControlText('E_GENERAL')= '' then
    begin
    SetControlEnabled('TGENE', false ) ;
    SetControlEnabled('GENE', false ) ;
    SetControlText('GENE', '' ) ;
    end
  else
    begin
    SetControlEnabled('TGENE', true ) ;
    SetControlEnabled('GENE', true ) ;
    NomTable := '' ;
    NatureCompte := RechNatureCpte ('G',GetControlText('E_GENERAL')) ;
    if NatureCompte <> AncNature then SetControlText('GENE', '' ) ;
    if NatureCompte = 'BQE' then NomTable := 'TZGBANQUE' ;
    if NatureCompte = 'CAI' then NomTable := 'TZGCAISSE' ;
    if NatureCompte = 'CHA' then NomTable := 'TZGCHARGE' ;
    if NatureCompte = 'COC' then NomTable := 'TZGCOLLCLIENT' ;
    if NatureCompte = 'COD' then NomTable := 'TZGCOLLDIVERS' ;
    if NatureCompte = 'COF' then NomTable := 'TZGCOLLFOURN' ;
    if NatureCompte = 'COS' then NomTable := 'TZGCOLLSALARIE' ;
    if NatureCompte = 'DIV' then NomTable := 'TZGDIVERS' ;
    if NatureCompte = 'EXT' then NomTable := 'TZGEXTRA' ;
    if NatureCompte = 'IMO' then NomTable := 'TZGIMMO' ;
    if NatureCompte = 'PRO' then NomTable := 'TZGPRODUIT' ;
    if Copy(NatureCompte,1,2) = 'TI' then NomTable := 'TZGTIDTIC' ;
    SetControlProperty('GENE', 'DATATYPE', NomTable ) ;
    end ;
end ;

procedure TOF_CPMODCPTEECR.EAuxiliaireOnChange (Sender : TObject ) ;
var Nature, NomTable : string ;
    AncTable : string ;
begin

  AncTable := THEdit(GetControl('AUXI')).Datatype ;
  if GetControlText('E_AUXILIAIRE') = '' then
    begin
    SetControlEnabled('TAUXI', false ) ;
    SetControlEnabled('AUXI', false ) ;
    SetControlText('AUXI', '' ) ;
    end
  else
    begin
    SetControlEnabled('TAUXI', true ) ;
    SetControlEnabled('AUXI', true ) ;
    Nature := RechNatureCpte ('T',GetControlText('E_AUXILIAIRE')) ;
    if Nature = 'CLI' then NomTable := 'TZTCLIENT' ;
    if Nature = 'AUC' then NomTable := 'TZTCREDITEUR' ;
    if Nature = 'AUD' then NomTable := 'TZTDEBITEUR' ;
    if Nature = 'DIV' then NomTable := 'TZTDIVERS' ;
    if Nature = 'FOU' then NomTable := 'TZTFOURN' ;
    if Nature = 'SAL' then NomTable := 'TZTSALARIE' ;
    SetControlProperty('AUXI','DATATYPE',NomTable) ;
    if NomTable <> AncTable then SetControlText('AUXI', '' ) ;
    end ;
end ;

procedure TOF_CPMODCPTEECR.ToutEcrOnClick (Sender : TObject ) ;
begin
  if GetControlText('TOUTECR')='X' then
    begin
    Exercice.Value := '';
    DateComptable.Text:=stDate1900 ; DateComptable_.Text:=stDate2099 ;
    SetControlText('E_NUMEROPIECE', '' ) ;
    SetControlText('E_NUMEROPIECE_', '' ) ;
    THValComboBox(GetControl('E_QUALIFPIECE')).itemIndex := 0 ;
    SetControlEnabled('PCOMPLEMENT', false ) ;
    end
  else
    begin
    SetControlEnabled('PCOMPLEMENT', true ) ;
    end ;
end ;

procedure TOF_CPMODCPTEECR.ExerciceOnChange (Sender : TObject ) ;
begin
  ExoToDates(Exercice.Value,DateComptable,DateComptable_) ;
  if Exercice.Value='' then BEGIN DateComptable.Text:=stDate1900 ; DateComptable_.Text:=stDate2099 ; END ;
end ;

procedure TOF_CPMODCPTEECR.PrechargeOrigines (ModifLettrage : boolean ) ;
Var Q : TQuery ;
    ReqSql, St, WhereLet : String ;
    TobAna : Tob ;
    idxEcr, p : integer ;
begin
  // recherche des écritures
  ReqSql := 'SELECT E_JOURNAL,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE,E_NUMECHE,E_QUALIFPIECE,' ;
  ReqSql := ReqSql + 'E_GENERAL,E_AUXILIAIRE,E_ANA,E_CREDIT,E_CREDITDEV,E_CREDITEURO,E_DEBIT,E_DEBITDEV,E_DEBITEURO,' ;
  ReqSql := ReqSql + 'E_REFINTERNE,E_LIBELLE,E_NATUREPIECE,E_VALIDE,E_ETABLISSEMENT,E_DEVISE,E_TAUXDEV,E_ETAT,' ;
  ReqSql := ReqSql + 'E_DATETAUXDEV,E_ECRANOUVEAU,E_SAISIEEURO,E_CONTREPARTIEGEN,E_CONTREPARTIEAUX,E_PERIODE,E_SEMAINE,' ;
  ReqSql := ReqSql + 'E_DATEPAQUETMAX,E_DATEPAQUETMIN,E_DATEPOINTAGE,E_PAQUETREVISION,E_REFPOINTAGE FROM ECRITURE ' ;
  St:=AnsiUpperCase(RecupWhereCritere(Pages)) ;
  WhereLet := '' ;
  if not ModifLettrage then WhereLet := ' AND (E_ETATLETTRAGE<>"PL" AND E_ETATLETTRAGE<>"TL")' ;
  ReqSql := ReqSql + St + whereLet+ ' ORDER BY E_JOURNAL,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE ' ;
  Q := OpenSQL(ReqSql, True) ;
  TobEcr.LoadDetailDB('ECRITURE','','',Q,False,True) ;
  Ferme(Q) ;
  if TOBEcr.Detail.Count>0 then TOBEcr.Detail[0].AddChampSup('MARQUE',True) ;

  // Recherche des lignes analytiques
  If (TobEcr.Detail.Count > 0) Then
  begin
  ReqSql := 'SELECT Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_GENERAL, Y_AUXILIAIRE, Y_NUMEROPIECE, Y_NUMLIGNE, Y_QUALIFPIECE,' ;
  ReqSql := ReqSql + ' Y_AXE, Y_SECTION, Y_NUMVENTIL FROM ANALYTIQ ' ;
  p := Pos ('E_', St ) ;
  While p>0 do
    begin
    System.Delete(St,p,Length('E_')) ; System.Insert('Y_',St,p) ;
    p := Pos('E_', St ) ;
    end ;
  ReqSql := ReqSql + St ;
  ReqSql := ReqSql + ' ORDER BY Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE' ;
  Q:=OpenSql(ReqSql,True) ;
  If not Q.Eof then
    begin
    TobAna:=TOB.Create('_ANALYTIQ', Nil, -1) ;
    TobAna.LoadDetailDB('ANALYTIQ', '', '', Q, True, True) ;
    // mise à jour de la TobEcr avec la TobAna
    For IdxEcr := 0 to TobEcr.Detail.Count -1 do
      begin
      While (TobAna.Detail.Count > 0) and (TobEcr.Detail[IdxEcr].GetValue('E_NUMEROPIECE')=TobAna.Detail[0].GetValue('Y_NUMEROPIECE'))
         and (TobEcr.Detail[IdxEcr].GetValue('E_NUMLIGNE')=TobAna.Detail[0].GetValue('Y_NUMLIGNE'))
         and (TobEcr.Detail[IdxEcr].GetValue('E_GENERAL')=TobAna.Detail[0].GetValue('Y_GENERAL'))
         and (TobEcr.Detail[IdxEcr].GetValue('E_AUXILIAIRE')=TobAna.Detail[0].GetValue('Y_AUXILIAIRE')) Do
         begin
           TobAna.Detail[0].ChangeParent(TobEcr.Detail[IdxEcr], -1) ;
         end ;
      end ;
    TobAna.free ;
    end ;
  Ferme(Q) ;
  end ;

end ;

procedure TOF_CPMODCPTEECR.MarqueOrigine ;
var TobE : Tob ;
begin
  TobE := TobEcr.FindFirst(['E_JOURNAL','E_EXERCICE','E_DATECOMPTABLE','E_NUMEROPIECE','E_NUMLIGNE'],
                           [QListeEcr.FindField('E_JOURNAL').AsString,
                            QListeEcr.FindField('E_EXERCICE').AsString,
                            QListeEcr.FindField('E_DATECOMPTABLE').AsDateTime,
                            QListeEcr.FindField('E_NUMEROPIECE').AsInteger,
                            QListeEcr.FindField('E_NUMLIGNE').AsInteger],False) ;
  if TobE<>Nil then TobE.PutValue('MARQUE','X') ;
end ;

procedure TOF_CPMODCPTEECR.VireInutiles ;
var idxEcr : integer ;
    TobE : Tob ;
begin
  for idxEcr := TobEcr.Detail.Count -1 downto 0 do
    begin
    TobE := TobEcr.Detail[idxEcr] ;
    if TobE.GetValue('MARQUE') <> 'X' then begin TobE.Free ; TobE := Nil ; end ;
    end ;
end ;

procedure TOF_CPMODCPTEECR.RecupMultiSelect( ModifLettrage : boolean) ;
var idxEcr : integer ;
begin
  TobEcr.ClearDetail ;
  // récupération des comptes sélectionnés de l'écran
  PrechargeOrigines (ModifLettrage) ;
  if Not ListeEcr.AllSelected then
    begin
    for idxEcr:=0 to ListeEcr.NbSelected-1 do
      begin
      ListeEcr.GotoLeBookmark(idxEcr) ;
      MarqueOrigine ;
      end ;
    VireInutiles ;
    end ;
end ;

function TOF_CPMODCPTEECR.VerifPointageCpt(CptOri, CptDest : string) : boolean ;
var Req, Where : string ;
    QOri, QDest : TQuery ;
    i : integer ;
    fb : TFichierBase ;
begin
  // recherche si le pointage est à mettre à jour et les modes de ventilation des comptes
  Req := 'SELECT G_POINTABLE, G_VENTILABLE, G_VENTILABLE1, G_VENTILABLE2, G_VENTILABLE3, G_VENTILABLE4, G_VENTILABLE5 FROM GENERAUX' ;
  Where := ' WHERE G_GENERAL="'+CptOri+'"' ;
  QOri := OpenSql(Req+Where,true) ;
  VentilOri := (QOri.FindField('G_VENTILABLE').asString = 'X') ;
  i := 1 ;
  For fb := fbAxe1 to fbAxe5 do
    begin
    VentilCptOri[fb] := QOri.FindField('G_VENTILABLE'+IntToStr(i)).asString;
    inc(i) ;
    end ;
  Where := ' WHERE G_GENERAL="'+CptDest+'"' ;
  QDest := OpenSql(Req+Where,true) ;
  VentilDes := (QDest.FindField('G_VENTILABLE').asString = 'X') ;
  i := 1 ;
  For fb := fbAxe1 to fbAxe5 do
    begin
    VentilCptDes[fb] := QDest.FindField('G_VENTILABLE'+intToStr(i)).asString;
    inc(i) ;
    end ;
  Result := (QOri.Fields[0].asString = 'X') and (QDest.Fields[0].asString <> 'X') ;
  Ferme(QOri) ;
  Ferme(QDest) ;
end ;

procedure TOF_CPMODCPTEECR.ChargeTobSection ( var TobSection : Tob ) ;
var Req : string ;
    Q : TQuery ;
begin
  TobSection := Tob.Create('LesSections', Nil, -1 ) ;
  Req := 'SELECT S_AXE, S_SECTION, S_TOTALDEBIT, S_TOTALCREDIT, S_TOTDEBE, S_TOTCREE, S_TOTDEBS, S_TOTCRES FROM SECTION ORDER BY S_AXE, S_SECTION' ;
  Q := OpenSQL(Req, False ) ;
  TobSection.LoadDetailDB('SECTION','','',Q,false ) ;
  if TobSection.Detail.Count>0 then TobSection.Detail[0].AddChampSup('MODIF',True) ;
  Ferme (Q) ;
end ;

function TOF_CPMODCPTEECR.TraiteCptBqe ( Compte : string ; var NewJal : string ) : boolean ;
var Q : TQuery ;
    Req : string ;
begin
  Req := 'SELECT J_JOURNAL FROM JOURNAL WHERE J_CONTREPARTIE="'+Compte+'"' ;
  Q := openSql(Req,true ) ;
  if not Q.eof then NewJal := Q.Fields[0].asString ;
  Ferme(Q) ;
  result := NewJal <> '' ;
end ;

procedure TOF_CPMODCPTEECR.ChargeTobJal ;
var Q : TQuery ;
    Req : string ;
begin
  TobJal := Tob.Create('MesJournaux',Nil,-1 ) ;
  Req := 'SELECT J_JOURNAL,J_TOTALCREDIT,J_TOTALDEBIT,J_TOTCREE,J_TOTCRES,J_TOTDEBE,J_TOTDEBS FROM JOURNAL' ;
  Q := OpenSql(Req, True ) ;
  TobJal.LoadDetailDB('JOURNAL','','',Q,false ) ;
  Ferme(Q) ;
  if TobJal.Detail.Count>0 then TobJal.Detail[0].AddChampSup('MODIF',True) ;
end ;

procedure TOF_CPMODCPTEECR.CalculCumul(Exo : string; Debit, Credit : double ) ;
begin
  CumDeb := CumDeb + Debit ;
  CumCre := CumCre + Credit ;
  if Exo = VH^.Precedent.Code then
    begin // cumul exercice précédent
    CumDebP := CumDebP + Debit ;
    CumCreP := CumCreP + Credit ;
    end
  else if Exo = VH^.EnCours.Code then
         begin // cumul exercice en cours
         CumDebE := CumDebE + Debit ;
         CumCreE := CumCreE + Credit ;
         end
       else
         begin // Cumul exercice suivant
         CumDebS := CumDebS + Debit ;
         CumCreS := CumCreS + Credit ;
         end ;
end ;

procedure TOF_CPMODCPTEECR.ModifLigneAnaAuxi (TobE : Tob ; CptAuxi : String ) ;
var idxAna : integer ;
    TobAna : Tob ;
begin
  for idxAna := 0 to TobE.Detail.Count-1 do
    begin
    TobAna := TobE.Detail[idxAna] ;
    TobAna.PutValue('Y_AUXILIAIRE',CptAuxi ) ;
    end ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 13/12/2001
Modifié le ... : 13/12/2001
Description .. : Mise à jour des cumuls dans les tables soit pour comptes 
Suite ........ : généraux, soit
Suite ........ : pour tiers, soit pour journaux.
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMODCPTEECR.MajSolde( P, Signe : char; Code : string ) ;
var NomTable, NomCol, Req : string ;
begin
  Case P of
    'G' : begin NomTable := 'GENERAUX' ; NomCol := 'G_GENERAL' ; end ;
    'T' : begin NomTable := 'TIERS' ; NomCol := 'T_AUXILIAIRE' ; end ;
    'J' : begin NomTable := 'JOURNAL' ; NomCol := 'J_JOURNAL' ; end ;
  end ;
  Req := 'UPDATE '+NomTable+' SET '+P+'_TOTALDEBIT='+P+'_TOTALDEBIT '+Signe+' '+StrFPoint(CumDeb)+' ,' ;
  Req := Req +P+'_TOTALCREDIT='+P+'_TOTALCREDIT '+Signe+' '+StrFPoint(CumCre)+' ,' ;
  If P = 'T' then
    begin
    Req := Req +P+'_TOTCREP='+P+'_TOTCREP '+Signe+' '+StrFPoint(CumCreP)+' ,' ;
    Req := Req +P+'_TOTDEBP='+P+'_TOTDEBP '+Signe+' '+StrFPoint(CumDebP)+' ,' ;
    end ;
  Req := Req +P+'_TOTCREE='+P+'_TOTCREE '+Signe+' '+StrFPoint(CumCreE)+' ,' ;
  Req := Req +P+'_TOTDEBE='+P+'_TOTDEBE '+Signe+' '+StrFPoint(CumDebE)+' ,' ;
  Req := Req +P+'_TOTCRES='+P+'_TOTCRES '+Signe+' '+StrFPoint(CumCreS)+' ,' ;
  Req := Req +P+'_TOTDEBS='+P+'_TOTDEBS '+Signe+' '+StrFPoint(CumDebS) ;
  Req := Req + ' WHERE '+NomCol+'="'+Code+'"' ;
  ExecuteSql(Req) ;
end ;

procedure TOF_CPMODCPTEECR.ModifJal(TobE : Tob ) ;
var TobJ : tob ;
begin
  // mise à jour des cumuls du journal
  TobJ := TobJal.FindFirst(['J_JOURNAL'],[TobE.GetValue('E_JOURNAL')],False) ;
  if TobJ<>Nil then
    begin
    TobJ.PutValue('MODIF','X') ;
    TobJ.PutValue('J_TOTALDEBIT', TobJ.GetValue('J_TOTALDEBIT')-TobE.GetValue('E_DEBIT')) ;
    TobJ.PutValue('J_TOTALCREDIT', TobJ.GetValue('J_TOTALCREDIT')-TobE.GetValue('E_CREDIT')) ;
    if TobE.GetValue('E_EXERCICE') = VH^.EnCours.Code then
      begin
      TobJ.PutValue('J_TOTDEBE', TobJ.GetValue('J_TOTDEBE')-TobE.GetValue('E_DEBIT')) ;
      TobJ.PutValue('J_TOTCREE', TobJ.GetValue('J_TOTCREE')-TobE.GetValue('E_CREDIT')) ;
      end
    else
      begin
      TobJ.PutValue('J_TOTDEBS', TobJ.GetValue('J_TOTDEBS')-TobE.GetValue('E_DEBIT')) ;
      TobJ.PutValue('J_TOTCRES', TobJ.GetValue('J_TOTCRES')-TobE.GetValue('E_CREDIT')) ;
      end ;
    end ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 13/12/2001
Modifié le ... :   /  /    
Description .. : Suppression de la tob des journaux non modifiés.
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMODCPTEECR.VireJal ;
var idxJal : integer ;
    TobJ : Tob ;
begin
  for idxJal := TobJal.Detail.Count -1 downto 0 do
    begin
    TobJ := TobJal.Detail[idxJal] ;
    if TobJ.GetValue('MODIF') <> 'X' then begin TobJ.Free ; TobJ := Nil ; end ;
    end ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 13/12/2001
Modifié le ... :   /  /    
Description .. : Suppression de la tob des sections non modifiées.
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMODCPTEECR.VireSection (TobSection : tob ) ;
var idxSec : integer ;
    TobS : Tob ;
begin
  for idxSec := TobSection.Detail.Count -1 downto 0 do
    begin
    TobS := TobSection.Detail[idxSec] ;
    if TobS.GetValue('MODIF') <> 'X' then begin TobS.Free ; TobS := Nil ; end ;
    end ;
end ;

procedure TOF_CPMODCPTEECR.MajSection ( TobS: Tob; Exo : String; TobA : Tob; Signe : integer ) ;
var TobSec : Tob ;
begin
  TobSec := TobS.FindFirst(['S_AXE','S_SECTION'],[TobA.GetValue('Y_AXE'),TobA.GetValue('Y_SECTION')],False) ;
  if TobSec<>Nil then
    begin
    TobSec.PutValue('MODIF','X') ;
    TobSec.PutValue('S_TOTALDEBIT', TobSec.GetValue('S_TOTALDEBIT')+(Signe * TobA.GetValue('Y_DEBIT'))) ;
    TobSec.PutValue('S_TOTALCREDIT', TobSec.GetValue('S_TOTALCREDIT')+(Signe * TobA.GetValue('Y_CREDIT'))) ;
    if Exo = VH^.EnCours.Code then
      begin
      TobSec.PutValue('S_TOTDEBE', TobSec.GetValue('S_TOTDEBE')+(Signe * TobA.GetValue('Y_DEBIT'))) ;
      TobSec.PutValue('S_TOTCREE', TobSec.GetValue('S_TOTCREE')+(Signe * TobA.GetValue('Y_CREDIT'))) ;
      end
    else
      begin
      TobSec.PutValue('S_TOTDEBS', TobSec.GetValue('S_TOTDEBS')+(Signe * TobA.GetValue('Y_DEBIT'))) ;
      TobSec.PutValue('S_TOTCRES', TobSec.GetValue('S_TOTCRES')+(Signe * TobA.GetValue('Y_CREDIT'))) ;
      end ;
    end ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 13/12/2001
Modifié le ... :   /  /    
Description .. : Gestion des lignes analytiques en cas de changement de 
Suite ........ : compte général : 
Suite ........ : - si compte destination non ventilable : suppression des
Suite ........ : lignes.
Suite ........ : - si compte destination ventilable sur un autre axe : création 
Suite ........ : d'une ligne à 100% sur section attente.
Suite ........ : - si compte destination ventilable de la même façon : 
Suite ........ : modification des lignes.
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMODCPTEECR.TraitementAnaGene( TobE, TobSection : Tob ) ;
var fb : TFichierBase ;
    idxAna :integer ;
    TobA : Tob ;
begin
  for fb := fbAxe1 to fbAxe5 do
    begin
    if (VentilCptOri[fb]='X') and (VentilCptDes[fb]='-') then
      begin
      // suppression des lignes analytiques
      For idxAna := TobE.Detail.count-1 downto 0 do
        begin
        TobA := TobE.Detail[idxAna] ;
        MajSection (TobSection, TobE.GetValue('E_EXERCICE'), TobA, -1) ;
        TobA.DeleteDB ;
        end ;
      end ;
    if (VentilCptOri[fb]='X') and (VentilCptDes[fb]='X') then
      begin
      // modification des lignes analytiques
      For idxAna := TobE.Detail.count-1 downto 0 do
        begin
        TobA := TobE.Detail[idxAna] ;
        TobA.PutValue('Y_GENERAL', GetControlText('GENE')) ;
        end ;
      end ;
    if (VentilCptOri[fb]='-') and (VentilCptDes[fb]='X') then
      begin
      // création de la ligne 100% sur section d'attente
      TobA := Tob.Create('ANALYTIQ', Nil,-1 ) ;
      TobA.AddChampSupValeur('Y_GENERAL', GetControlText('GENE')) ;
      TobA.AddChampSupValeur('Y_AXE', fbToAxe(fb) ) ;
      TobA.AddChampSupValeur('Y_DATECOMPTABLE', TobE.GetValue('E_DATECOMPTABLE')) ;
      TobA.AddChampSupValeur('Y_NUMEROPIECE', TobE.GetValue('E_NUMEROPIECE')) ;
      TobA.AddChampSupValeur('Y_NUMLIGNE', TobE.GetValue('E_NUMLIGNE')) ;
      TobA.AddChampSupValeur('Y_SECTION', VH^.cpta[fb].Attente ) ;
      TobA.AddChampSupValeur('Y_EXERCICE', TobE.GetValue('E_EXERCICE')) ;
      TobA.AddChampSupValeur('Y_DEBIT', TobE.GetValue('E_DEBIT')) ;
      TobA.AddChampSupValeur('Y_CREDIT', TobE.GetValue('E_CREDIT')) ;
      TobA.AddChampSupValeur('Y_REFINTERNE', TobE.GetValue('E_REFINTERNE')) ;
      TobA.AddChampSupValeur('Y_LIBELLE', TobE.GetValue('E_LIBELLE')) ;
      TobA.AddChampSupValeur('Y_NATUREPIECE', TobE.GetValue('E_NATUREPIECE')) ;
      TobA.AddChampSupValeur('Y_QUALIFPIECE', TobE.GetValue('E_QUALIFPIECE')) ;
      TobA.AddChampSupValeur('Y_TYPEANALYTIQUE', 'N') ;
      TobA.AddChampSupValeur('Y_VALIDE', TobE.GetValue('E_VALIDE')) ;
      TobA.AddChampSupValeur('Y_ETAT', TobE.GetValue('E_ETAT')) ;
      TobA.AddChampSupValeur('Y_ETABLISSEMENT', TobE.GetValue('E_ETABLISSEMENT')) ;
      TobA.AddChampSupValeur('Y_DEBITEURO', TobE.GetValue('E_DEBITEURO')) ;
      TobA.AddChampSupValeur('Y_CREDITEURO', TobE.GetValue('E_CREDITEURO')) ;
      TobA.AddChampSupValeur('Y_DEVISE', TobE.GetValue('E_DEVISE')) ;
      TobA.AddChampSupValeur('Y_DEBITDEV', TobE.GetValue('E_DEBITDEV')) ;
      TobA.AddChampSupValeur('Y_CREDITDEV', TobE.GetValue('E_CREDITDEV')) ;
      TobA.AddChampSupValeur('Y_TAUXDEV', TobE.GetValue('E_TAUXDEV')) ;
      TobA.AddChampSupValeur('Y_DATETAUXDEV', TobE.GetValue('E_DATETAUXDEV')) ;
      TobA.AddChampSupValeur('Y_POURCENTAGE', 100 ) ;
      if TobE.GetValue('E_DEBIT') <> 0 Then
        begin
        TobA.AddChampSupValeur('Y_TOTALECRITURE', TobE.GetValue('E_DEBIT')) ;
        TobA.AddChampSupValeur('Y_TOTALDEVISE', TobE.GetValue('E_DEBITDEV')) ;
        TobA.AddChampSupValeur('Y_TOTALEURO', TobE.GetValue('E_DEBITEURO')) ;
        end
      else
        begin
        TobA.AddChampSupValeur('Y_TOTALECRITURE', TobE.GetValue('E_CREDIT')) ;
        TobA.AddChampSupValeur('Y_TOTALDEVISE', TobE.GetValue('E_CREDITDEV')) ;
        TobA.AddChampSupValeur('Y_TOTALEURO', TobE.GetValue('E_CREDITEURO')) ;
        end ;
      TobA.AddChampSupValeur('Y_JOURNAL', TobE.GetValue('E_JOURNAL')) ;
      TobA.AddChampSupValeur('Y_NUMVENTIL', 1 ) ;
      TobA.AddChampSupValeur('Y_ECRANOUVEAU', TobE.GetValue('E_ECRANOUVEAU')) ;
      TobA.AddChampSupValeur('Y_SAISIEEURO', TobE.GetValue('E_SAISIEEURO')) ;
      TobA.AddChampSupValeur('Y_CONTREPARTIEGEN', TobE.GetValue('E_CONTREPARTIEGEN')) ;
      TobA.AddChampSupValeur('Y_CONTREPARTIEAUX', TobE.GetValue('E_CONTREPARTIEAUX')) ;
      TobA.AddChampSupValeur('Y_AUXILIAIRE', TobE.GetValue('E_AUXILIAIRE')) ;
      TobA.AddChampSupValeur('Y_PERIODE', TobE.GetValue('E_PERIODE')) ;
      TobA.AddChampSupValeur('Y_SEMAINE', TobE.GetValue('E_SEMAINE')) ;
      TobA.InsertDB(Nil) ;
      MajSection (TobSection, TobE.GetValue('E_EXERCICE'), TobA, 1) ;
      TobA.Free ;
      end ;
    end ;
end ;

procedure TOF_CPMODCPTEECR.RechContrepartie(var TobContre, TobAnaContre : tob;TobE  :tob; NatCpte : string ) ;
var Req : string ;
    Q : TQuery ;
    NomCol1, NomCol2 : string ;
begin
  // recherche des lignes d'écritures pour mettre à jour les comptes de contrepartie
  NomCol1 := 'CONTREPARTIEGEN' ; NomCol2 := 'GENERAL' ;
  if NatCpte = 'T' then begin NomCol1 := 'CONTREPARTIEAUX' ; NomCol2 := 'AUXILIAIRE' ; end ;
  Req := 'SELECT E_JOURNAL,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE,E_NUMECHE,E_QUALIFPIECE,' ;
  Req := Req + 'E_PAQUETREVISION,E_'+NomCol1+' FROM ECRITURE ' ;
  Req := Req + 'WHERE E_JOURNAL="'+TobE.GetValue('E_JOURNAL')+'" AND E_EXERCICE="'+TobE.GetValue('E_EXERCICE')+'"' ;
  Req := Req + ' AND E_DATECOMPTABLE="'+USDATETIME(TobE.GetValue('E_DATECOMPTABLE'))+'" AND E_NUMEROPIECE='+StrFPoint(TobE.GetValue('E_NUMEROPIECE')) ;
  Req := Req + ' AND E_QUALIFPIECE="'+TobE.GetValue('E_QUALIFPIECE')+'" AND E_'+NomCol1+'="'+TobE.GetValue('E_'+NomCol2)+'"' ;
  Q := OpenSql(Req,false) ;
  TobContre.LoadDetailDB('ECRITURE','','',Q,true ) ;
  Ferme(Q) ;

  // recherche des lignes analytiques
  Req := 'SELECT Y_JOURNAL,Y_EXERCICE,Y_DATECOMPTABLE,Y_NUMEROPIECE,Y_NUMLIGNE,Y_AXE,Y_NUMVENTIL,Y_QUALIFPIECE,' ;
  Req := Req + 'Y_'+NomCol1+' FROM ANALYTIQ ' ;
  Req := Req + 'WHERE Y_JOURNAL="'+TobE.GetValue('E_JOURNAL')+'" AND Y_EXERCICE="'+TobE.GetValue('E_EXERCICE')+'"' ;
  Req := Req + ' AND Y_DATECOMPTABLE="'+USDATETIME(TobE.GetValue('E_DATECOMPTABLE'))+'" AND Y_NUMEROPIECE='+StrFPoint(TobE.GetValue('E_NUMEROPIECE')) ;
  Req := Req + ' AND Y_QUALIFPIECE="'+TobE.GetValue('E_QUALIFPIECE')+'" AND Y_'+NomCol1+'="'+TobE.GetValue('E_'+NomCol2)+'"' ;
  Q := OpenSql(Req,false) ;
  TobAnaContre.LoadDetailDB('ANALYTIQ','','',Q,true ) ;
  Ferme(Q) ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 13/12/2001
Modifié le ... :   /  /    
Description .. : mise à jour des comptes de contrepartie.
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMODCPTEECR.MajTobContre (var TobContre, TobAnaContre : tob ; NatCpte, NewCpte : string ) ;
var idxTob : integer ;
    NomCol : string ;
begin
  NomCol := 'CONTREPARTIEGEN' ; if NatCpte = 'T' then NomCol := 'CONTREPARTIEAUX' ;
  For idxTob := 0 to TobContre.detail.count-1 do
    begin
    TobContre.Detail[idxTob].PutValue('E_'+NomCol, NewCpte ) ;
    TobContre.Detail[idxTob].putValue('E_PAQUETREVISION', 1) ;
    end ;
  For idxTob := 0 to TobAnaContre.detail.count-1 do TobAnaContre.Detail[idxTob].PutValue('Y_'+NomCol, NewCpte ) ;
end ;

procedure TOF_CPMODCPTEECR.MajJournal ( CodeJal : string ) ;
begin
  VireJal ;
  TobJal.UpdateDB ;
  MajSolde ('J', '+' , CodeJal ) ;
  TobJal.Free ;
end ;

function TOF_CPMODCPTEECR.TraitementEcriture(ModifGene, ModifLettrage: boolean) : boolean ;
var NewJal : string ;
    ModJal : boolean ;
    MajPointage : boolean ;
    idxEcr : integer ;
    TobE, TobSection : Tob ;
    TobContre, TobAnaContre : Tob ;
begin
  result := false ;
  CumDeb:= 0 ; CumCre := 0 ; CumDebP := 0 ; CumCreP := 0 ; CumDebE := 0 ; CumCreE := 0 ; CumDebS := 0 ; CumCreS := 0 ;
  NewJal := '' ;
  MajPointage := false ; ModJal := False ;
  if NatureCompte = 'BQE' then ModJal := TraiteCptBqe(GetControlText('GENE'),NewJal) ;
  if ModJal then
    begin
    if PGIAsk(TraduireMemoire('Le journal des écritures va être modifié en '+NewJal+'. Confirmez-vous le traitement ?'),TraduireMemoire(Titre)) = mrNo then exit ;
    ChargeTobJal ;
    end ;

  TobContre := Tob.Create('Contrepartie', Nil, -1 ) ;
  TobAnaContre := Tob.Create('Anacontrepartie', Nil, -1 ) ;
  if ModifGene then MajPointage := VerifPointageCpt (GetControlText('E_GENERAL'),GetControlText('GENE')) ;
  if ModifGene and (VentilOri or VentilDes) then ChargeTobSection(TobSection) ;

  Try
    BeginTrans ;
    For idxEcr := 0 to TobEcr.Detail.Count-1 do
      begin
      TobE := TobEcr.Detail[idxEcr] ;
      // calcul des cumuls crédit / débit
      CalculCumul(TobE.GetValue('E_EXERCICE'),TobE.GetValue('E_DEBIT'),TobE.GetValue('E_CREDIT')) ;
      TobE.PutValue('E_PAQUETREVISION', 1) ;

      if ModifGene then
        begin
        // traitement de l'analytique
        if (VentilOri or VentilDes) then
          begin
          if VentilDes then TobE.PutValue('E_ANA', 'X') else TobE.PutValue('E_ANA', '-') ;
          TraitementAnaGene ( TobE, TobSection ) ;
          end ;
        // recherche des autres écritures de la pièce pour mettre à jour le compte de contrepatrtie
        RechContrepartie (TobContre,TobAnaContre,TobE,'G' ) ;
        TobE.PutValue('E_GENERAL', GetControlText('GENE')) ;
        end
      else
        begin
        // traitement de l'analytique
        if TobE.GetValue('E_ANA')='X' then ModifLigneAnaAuxi(TobE, GetControlText('AUXI')) ;
        // recherche des autres écritures de la pièce pour mettre à jour le compte de contrepatrtie
        RechContrepartie (TobContre,TobAnaContre,TobE,'T' ) ;
        TobE.PutValue('E_AUXILIAIRE', GetControlText('AUXI')) ;
        end ;

      // modif pointage si besoin
      if MajPointage then
        begin
        TobE.PutValue('E_REFPOINTAGE', '' ) ;
        TobE.PutValue('E_DATEPOINTAGE', StrToDateTime(stDate1900)) ;
        end ;
      // modif du journal si besoin
      if ModJal then
        begin
        ModifJal(TobE) ;
        TobE.PutValue('E_JOURNAL', NewJal ) ;
        end ;
      end ;

    // mise à jour des écritures
    TobEcr.UpdateDB ;

    // mise à jour des comptes de contrepartie
    if ModifGene then MajTobContre (TobContre, TobAnaContre, 'G', GetControlText('GENE'))
                 else MajTobContre (TobContre, TobAnaContre, 'T', GetControlText('AUXI')) ;
    TobContre.UpdateDB ; TobAnaContre.UpdateDB ;
    TobContre.Free ; TobAnaContre.Free ;

    // mise à jour des journaux si besoin
    if ModJal then MajJournal(NewJal) ;

    // mise à jour des sections modifiées
    if TobSection <> Nil then
      begin
      VireSection(TobSection) ;
      TobSection.UpdateDB ;
      TobSection.Free ;
      end ;
    // mise à jour des cumuls des comptes modifiés
    if ModifGene then
      begin
      MajSolde ('G','-', GetControlText('E_GENERAL')) ;
      MajSolde ('G','+', GetControlText('GENE')) ;
      end
    else
      begin
      MajSolde ('T','-', GetControlText('E_AUXILIAIRE')) ;
      MajSolde ('T','+', GetControlText('AUXI')) ;
      end ;
    CommitTrans ;
  Except
    RollBack ;
    PGIBox(TraduireMemoire('Le traitement a échoué.'),TraduireMemoire(Titre)) ;
    Exit ;
  end ;
  result := true ;
end ;

procedure TOF_CPMODCPTEECR.BOuvrirOnClick (sender : TObject) ;
var TraitOk : boolean ;
    ModifGene, ModifLettrage : boolean ;
begin
  ListeEcr := Fiche.FListe ;
  QListeEcr := Fiche.Q ;

  // vérification qu'au moins une écriture est sélectionnée
  if (ListeEcr.nbSelected = 0) and (not ListeEcr.AllSelected) then
    begin
    PGIInfo(TraduireMemoire('Vous n''avez sélectionné aucune écriture. Traitement impossible.'), TraduireMemoire(Titre)) ;
    Exit ;
    end ;
  // le nouveau compte doit être renseigné
  if (GetControlText('GENE') = '') and (GetControlText('AUXI')='') then
    begin
    PGIInfo(TraduireMemoire('Vous devez renseigner un nouveau compte.'), TraduireMemoire(Titre)) ;
    Case THRadioGroup(GetControl('CHOIXCPTE')).ItemIndex of
      0 : SetFocusControl('GENE') ;
      1 : SetFocusControl('AUXI') ;
    end ;
    Exit ;
    end ;

  // Ne pas lancer le traitement si les comptes origine et destination sont identiques
  if (GetControlText('GENE') <> '') and (GetControlText('GENE') = GetControlText('E_GENERAL')) then
    begin
    PGIInfo(TraduireMemoire('Le compte de destination est identique à celui d''origine. Aucun traitement à faire.'), TraduireMemoire(Titre)) ;
    SetFocusControl('GENE') ;
    Exit ;
    end ;
  if (GetControlText('AUXI') <> '') and (GetControlText('AUXI') = GetControlText('E_AUXILIAIRE')) then
    begin
    PGIInfo(TraduireMemoire('Le compte de destination est identique à celui d''origine. Aucun traitement à faire.'), TraduireMemoire(Titre)) ;
    SetFocusControl('AUXI') ;
    Exit ;
    end ;
  If PGIASK(TraduireMemoire('Confirmez-vous le traitement ?'),TraduireMemoire(Titre)) = MrNo then exit ;

  ModifGene := THRadioGroup(GetControl('CHOIXCPTE')).ItemIndex = 0 ;

  if ModifGene then ModifLettrage := false
  else ModifLettrage := (GetControlText('TOUTECR')='X') and (ListeEcr.AllSelected) ;
  if Not ModifLettrage then PGIInfo(TraduireMemoire('Les écritures partiellement ou totalement lettrées ne seront pas modifiées.'), TraduireMemoire(Titre)) ;

  EnableControls(Fiche, False) ;
  // récupération des écritures sélectionnées
  RecupMultiSelect(ModifLettrage) ;
  // lancement du traitement
  TraitOk := TraitementEcriture(ModifGene, ModifLettrage) ;
  if TraitOK then PGIInfo(TraduireMemoire('Traitement terminé.'), TraduireMemoire(Titre)) ;
  EnableControls(Fiche, True) ;

  // rafraîchissement de la liste
  TToolBarButton97(GetControl('BCHERCHE')).Click ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 19/12/2001
Modifié le ... :   /  /    
Description .. : Visualisation de la pièce sur le double clic dans la liste des 
Suite ........ : écritures.
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPMODCPTEECR.FListeOnDblClick (sender : TObject) ;
begin
  TrouveEtLanceSaisie(QListeEcr, taConsult,GetControlText('E_QUALIFPIECE')) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 13/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMODCPTEECR.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;          

procedure TOF_CPMODCPTEECR.OnArgument (S : String ) ;
begin
  Inherited ;
  Fiche := TFMul(Ecran) ;
  ListeEcr := Fiche.FListe ;
  QListeEcr := Fiche.Q ;
  Pages := TPageControl(GetControl('PAGES')) ;

  THRadioGroup(GetControl('CHOIXCPTE')).OnClick := ChoixCpteOnClick ;
  THEdit(GetControl('E_GENERAL')).OnChange := EGeneralOnChange ;
  THEdit(GetControl('E_AUXILIAIRE')).OnChange := EAuxiliaireOnChange ;
  THValComboBox(GetControl('E_QUALIFPIECE')).ItemIndex := 0 ;
  TCheckBox(GetControl('TOUTECR')).OnClick := ToutEcrOnClick ;

  Exercice := THValComboBox(GetControl('E_EXERCICE')) ;
  DateComptable := THEdit(GetControl('E_DATECOMPTABLE')) ;
  DateComptable_ := THEdit(GetControl('E_DATECOMPTABLE_')) ;
  Exercice.ItemIndex := 0 ;
  DateComptable.Text:=stDate1900 ; DateComptable_.Text:=stDate2099 ;
  Exercice.onChange := ExerciceOnChange ;

  TobEcr := Tob.Create('MesEcritures', Nil, -1 ) ;
  TToolBarButton97(GetControl('BOUVRIR')).OnClick := BOuvrirOnClick ;
  ListeEcr.OnDblClick := FListeOnDblClick ;

  if GetParamSocSecur('SO_CPMULTIERS', false) then
    THEdit(GetControl('E_AUXILIAIRE', true)).OnElipsisClick:=AuxiElipsisClick;

end ;

procedure TOF_CPMODCPTEECR.OnClose ;
begin
  TobEcr.Free ;
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_CPMODCPTEECR ] ) ;
end.
