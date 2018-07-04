unit ImGenEcr;

// CA  - 17/05/1999 - supression des infos liées à l'indice et la couleur
// CA  - 07/07/1999 - IsCompteVentilable : fermeture de la Query
// CA  - 07/07/1999 - Pas de génération de dérogatoire si CumulCesFisc = 0
// CA  - 09/07/1999 - Ecritures de cession si CumulCesEco <>0
// CA  - 31/10/2000 - Calcul de la valeur de l'immo avant cession (cas du rapprochement).
// MBO - 24/10/2005 - crc 2002/10 - calculer l'exceptionnel en y incluant la dépreciation d'actif
// BTY - 10/05      - En CRC 2002-10, amortissement dérogatoire supplémentaire si modif base
// TGA - 07/11/2005 - pas de dépréciation dans l'exceptionnel
// TGA - 16/11/2005 - Réaffectation du montant de l'écriture au débit ou crédit selon son signe
// BTY - 11/05      - Si ajout d'un plan fiscal CRC 2002-10 pur en révision du plan d'amortissement,
//                      => dérogatoire supplémentaire
// MBO - 25/11/2005 - fq 17084 - type exceptionnel sur sortie non testé (RDO ou DOT)
// MBO - 25/11/2005 - fq 16400 - incidence de la prise en compte de l'exceptionnel sur sortie
// PGR - 12/2005    - Opération de Changement de méthode CDA dans l'exercice
// PGR - 12/2005    - Opération de Changement de méthode (Eclatement préalable) ACQ dans l'exercice
// TGA - 21/12/2005 - FQ 17215 - GetParamSoc => GetParamSocSecur
// BTY - 02/06      - FQ 15609 En niveau détaillé, libellé écriture = celui de l'immo au lieu du libellé compte
//                    et Réf pièce = référence facture immo au lieu du code immo
// BTY - 03/06 -      FQ 17713 Si un compte ventilable sans ventilations présente une dotation < 0
//                             on a un carton Indice hors limite(-1) dans ChargeVentilImo  cf dossier client MLOPTIONS
// BTY - 03/06 -      FQ 17748 - Ecritures d'intégration non regroupées si dotations négatives
// BTY - 03/06 -      FQ 17760 - Ecriture de dotation n'a pas le bon libellé si dotation négative
// MBO - 10/04/2006 - FQ 17773 - en rapprochement bancaire on ne génére aucun écart car on fait le rapprochement bancaire
//                               après avoir intégrées les écritures
// MBO - 14/04/2006 - FQ 16750 - en rapprochement bancaire, voir les comptes mvtés en cpta même si non mvtés immo
// MVG - 22/06/2006 - FQ 10319 - en rapprochement avec la compta, ne pas proratiser les valeurs d'actif
//                               si rapprochement avant la sortie, faire apparaitre l'immo
// MBO - 23/06/2006 - FQ 18470 - saisie d'un exceptionnel + Cloture + cession : l'exceptionnel est doublé
//                               ds le compte d'amortissement cédé
// MBO - 25/08/2006 -          - correction d'un conseil de compil
// MBO - 03/11/2006 -          - Génération des écritures pour la subvention d'investissement
// MVG - 05/12/2006 - FQ 19282
// MBO - 03/2007    - FQ 17512 - suppression du module encadré par ifdef plus_tard
// MBO - 25/06/2007 - FQ 20841 - pas de calcul sur plan variable à date intermédiaire
// MBO - 29/10/2007 - FQ 21754 - Ajout de 2 paramètres en appel du PlanInfo.calcul
interface

uses
   classes
   ,HCtrls
   ,SysUtils
   ,HEnt1
   ,graphics
   ,HStatus
   ,ParamSoc
   {$IFDEF SERIE1}
   {$ELSE}
   {$IFDEF MODENT1}
   ,CPProcMetier
   {$ENDIF MODENT1}
   ,Ent1
   {$ENDIF}
   ,ImEnt
   {$IFDEF eAGLCLient}
   , uTOB
   {$ELSE}
     {$IFNDEF DBXPRESS},dbtables{$ELSE},uDbxDataSet{$ENDIF}
   {$ENDIF eAGLCLient}
   ;

type
  TAmNiveauDetail = (ndSynth, ndCompte, ndDetail);
  {$IFDEF PLUS_TARD}
  TAmListeEcriture = class
    private
      fListe          : TList;
      fNiveauDetail   : TAmNiveauDetail;
      fCompteRef      : string;
    public
      procedure AjouteLigne ( Compte : string; Debit, Credit : double);
    published
      NiveauDetail : TAmNiveauDetail read fNiveauDetail write fNiveauDetail;
      CompteRef : string reaf fCompteRef write fCompteRef;   //@@@
  end;
{$ENDIF}


  TParamEcr = class
    Journal : string;
    Libelle : string;
    Date : TDateTime;
    DateCalcul : TDateTime;
    DateAmort : TDateTime;
    DateFinAmort : TDateTime;
    bDotation : boolean;
    Methode : string;
    CompteRef : string; // Compte de référence (compte 2...)
    Etablissement : string;
    Lib_Immo : string;  // FQ 15609 libellé immo à prévoir
    Ref_Immo : string;  // FQ 15609 référence immo à prévoir
  end;
  {$IFDEF SERIE1}
  {$ELSE}
  TAna = class
    NumeroVentil : string;
    Section,Libelle : string;
    TauxMontant,TauxQte1,TauxQte2,Montant : double;
    public
      procedure Copy (ARecord : TAna);
  end;
  {$ENDIF}
  TLigneEcriture = class
    CodeImmo : string;
    Piece : string;
    Compte : string;
    CompteRef : string;
    Auxi : string;
    Libelle : string;
    Debit : double;
    Credit : double;
    Date : TDateTime;
    CodeJournal : string;
    {$IFDEF SERIE1}
    {$ELSE}
    Axes : Array[0..ImMaxAxe-1] of TList;
    {$ENDIF}
    Couleur : TColor;
    Etablissement : string;
  public
      procedure Copy (LigneEcr : TLigneEcriture);
  end;

procedure VideListeEcritures(L : TList);
procedure CalculEcrituresDotation ( var L : TList;Code : string;ParamEcr : TParamEcr;NivDetail : integer);
procedure CalculEcrituresEcheances ( var L : TList;Code : string;ParamEch,ParamPai : TParamEcr;NivDetail : integer; dtEchMin, dtEchMax : TDateTime;AvecDejaIntegre: boolean=false);
function RecalculMontant (Valeur : double; DateCalcul,DateAmort,DateFinAmort : TDateTime;bDotation : boolean;Methode : string; AvecSurAMort : boolean) : double;


{$IFDEF SERIE1}
{$ELSE}
//procedure ChargeVentilImo (AxesVentil : string;var Axes:Array of TList;Nat,Cpte : String;MontantVentil,MontantTotal : double) ;
function  SommeVentilations ( L : TList) : double;
{$ENDIF}

implementation

uses
     {$IFDEF MODENT1}
     CPTypeCons,
     {$ENDIF MODENT1}
      ImOuPlan,
      ImPlan,
      ImOutGen
      {$IFDEF SERIE1}
      {$ELSE}
      ,ZCompte
      ,ZTiers
      {$ENDIF}
      ,ImPlanInfo
      ;

var ListeCompte : TList;
    Dotation_V2 : boolean;
    fdotation : Boolean;
    fcompte : String;
    fprorata : boolean; // MVG 22/06/2006

{$IFDEF SERIE1}
{$ELSE}
////////////////////////////////////////////////////////////////////////////////
procedure ChargeVentilImo (AxesVentil : string;var Axes:Array of TList;Nat,Cpte : String;MontantVentil,MontantTotal : double) ;
Var ARecord,ARec2: TAna ;i,v,x : integer ; QVentil : TQuery;wAtt, Requete : string;
    TotalMontant : double;
    wAxe: Tlist ;
begin
for i:=0 to ImMaxAxe-1 do
  begin
  if Pos(InttoStr(i+1),AxesVentil)>0 then
    begin
    TotalMontant:=0.0;
    wAxe:=TList.Create;
    //Calcul la ventilation paramètrée
    Requete := 'SELECT VENTIL.*,S_LIBELLE FROM VENTIL, SECTION WHERE VENTIL.V_NATURE="'+Nat+IntToStr(i+1)+'" AND VENTIL.V_COMPTE="'+Cpte+'"'
          + ' AND S_SECTION=V_SECTION ORDER BY V_NATURE, V_COMPTE, V_NUMEROVENTIL' ;
    QVentil:=OpenSQL (Requete,True);
    while not QVentil.EOF do
      begin
      ARec2:=TAna.Create ;
      ARec2.NumeroVentil:=QVentil.FindField('V_NUMEROVENTIL').AsString ;
      ARec2.Section     :=QVentil.FindField('V_SECTION').AsString ;
      ARec2.Libelle     :=QVentil.FindField('S_LIBELLE').AsString ;
      ARec2.TauxMontant :=QVentil.FindField('V_TAUXMONTANT').AsFloat ;
      ARec2.TauxQte1    :=QVentil.FindField('V_TAUXQTE1').AsFloat ;
      ARec2.TauxQte2    :=QVentil.FindField('V_TAUXQTE2').AsFloat ;
      ARec2.Montant     :=Arrondi(((MontantVentil*ARec2.TauxMontant)/100),2) ;
      TotalMontant      :=TotalMontant+ARec2.Montant ;
      if ARec2.Montant<>0 then wAxe.Add(ARec2);
      QVentil.Next ;
      end;
    Ferme(QVentil) ;
    //proratise les montants si le montant la ventil est > 100%
    if MontantVentil<TotalMontant then
      begin
{      for v:=0 to wAxe.Count-1 do
        begin
        ARecord:=wAxe.Items[v];
        ARecord.Montant:=Arrondi((ARecord.Montant*TotalMontant)/MontantVentil,2);
        end;}
        { CA - 27/03/2003 - si montant < Total de la ventilation, on diminue la dernière ventilation.}
        { ancienne méthode ne fonctionnait pas toujours ! } 
        ARecord := wAxe.Items[wAxe.Count-1];
        ARecord.Montant := Arrondi ( ARecord.Montant - (TotalMontant-MontantVentil), V_PGI.OkDecV );
      end ;
    //Calcul la section d'attente
    if MontantVentil>TotalMontant then
      begin
      wAtt:=VHImmo^.Cpta[ImAxeToFb('A'+IntToStr(i+1))].Attente ;
      QVentil:=OpenSQL ('SELECT S_LIBELLE FROM SECTION WHERE S_SECTION="'+wAtt+'"',True);
      if not QVentil.Eof then
        begin
        ARec2:=TAna.Create ;
        ARec2.NumeroVentil:='' ; //QVentil.FindField('V_NUMEROVENTIL').AsString ;
        ARec2.Section     :=wAtt ;
        ARec2.Libelle     :=QVentil.FindField('S_LIBELLE').AsString ;
        ARec2.TauxMontant :=0 ;
        ARec2.TauxQte1    :=0 ;
        ARec2.TauxQte2    :=0 ;
        ARec2.Montant     :=MontantVentil-TotalMontant ;
        if ARec2.Montant<>0 then wAxe.Add(ARec2);
        end ;
      Ferme(QVentil);
      end ;
    //report des ventilations dans la liste des axes
    if wAxe.Count>0 then
      begin
      if Axes[i]=nil then Axes[i]:=TList.Create;
      //pour tous les axes calculés
      for x:=0 to WAxe.Count-1 do
        begin
        ARec2:=wAxe.Items[x] ;
        Arecord:=nil ;
        for v:=0 to Axes[i].Count-1  do
          begin
          ARecord:=Axes[i].Items[v];
          if (ARecord.Section=ARec2.Section) and (ARecord.NumeroVentil=ARecord.NumeroVentil) then break;
          Arecord:=nil ;
          end;
        if Arecord=nil then
          begin
          ARecord:=TAna.Create;
          Arecord.Copy(ARec2) ;
          Axes[i].Add(ARecord);
          end
        else
          begin
          ARecord.Montant:=ARecord.Montant+ARec2.Montant ;
          end ;
        end ;
      end ;
    wAxe.Free ;
    //recalcul des taux
    for v:=0 to Axes[i].Count-1  do
      begin
      ARecord:=Axes[i].Items[v];
      ARecord.TauxMontant:=Arrondi((ARecord.Montant*100)/MontantTotal,2) ;
      end;
    end ;
  end ;
end ;

////////////////////////////////////////////////////////////////////////////////
function SommeVentilations ( L : TList) : double;
var i : integer; ARecord : TAna;
begin
  Result:=0.0;
  if L.Count<=0 then exit;
  for i:=0 to L.Count-1 do
  begin
    ARecord:=L.Items[i];
    result :=result+ARecord.Montant;
  end;
  result:=Arrondi(Result,V_PGI.OKDecV);
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
function IsCompteVentilable (Compte : string) : string;
var i : integer;
    Q : TQuery;
begin
  Result := '00000';
  Q := OpenSQL ('SELECT G_VENTILABLE,G_VENTILABLE1,G_VENTILABLE2,G_VENTILABLE3,G_VENTILABLE4,G_VENTILABLE5 FROM GENERAUX WHERE G_GENERAL="'+Compte+'"',True);
  if not Q.Eof then
  begin
// Suppression suite rq CH : tous les comptes sont ventilables.
//    if (Q.FindField('G_NATUREGENE').AsString = 'CHA') or
//       (Q.FindField('G_NATUREGENE').AsString = 'PRO')then
//    begin
      if Q.FindField('G_VENTILABLE').AsString<>'X' then
      begin Ferme (Q);exit;end;
      for i:=1 to 5 do
        Result[i] := IntToStr(i*Ord((Q.Fields[i].AsString)='X'))[1];
//    end;
  end;
  Ferme (Q);
end;
////////////////////////////////////////////////////////////////////////////////
procedure VideListeCompte(L : TList) ;
Var i : integer ;
BEGIN
  if L=Nil then Exit ; if L.Count<=0 then Exit ;
  for i:=0 to L.Count-1 do
  BEGIN
    Dispose(L[i]) ; L[i]:=Nil ;
  END ;
  L.Clear ;
END ;


////////////////////////////////////////////////////////////////////////////////
function RecupereLibelle (Compte : string; var General : boolean) : string;
var Q : TQuery;
    ARecord : ^TDefCompte;
    i : integer;
begin
  result := '';
  if (Compte='') or (ListeCompte=nil) then exit;
  for i:=0 to ListeCompte.Count - 1 do
  begin
    ARecord := ListeCompte.Items[i];
    if ARecord^.Compte=Compte then
    begin
      General:=ARecord^.bGeneral;
      result :=ARecord^.Libelle;
    end;
  end;
  if result = '' then
  begin
    Q:=OpenSQL ('SELECT G_LIBELLE FROM GENERAUX WHERE G_GENERAL="'+Compte+'"',TRUE);
    if not Q.EOF then
    begin
      new(ARecord);
      ARecord^.Compte   :=Compte;
      ARecord^.Libelle  :=Q.FindField('G_LIBELLE').AsString;
      ARecord^.bGeneral :=True;
      General           :=True;
      result            :=ARecord^.Libelle ;
      ListeCompte.Add(ARecord);
    end ;
    ferme(Q) ;
  end ;
  if result = '' then
  begin
    Q:=OpenSQL ('SELECT T_LIBELLE FROM TIERS WHERE T_AUXILIAIRE="'+Compte+'"',TRUE);
    if not Q.EOF then
    begin
      new(ARecord);
      ARecord^.Compte  :=Compte;
      ARecord^.Libelle :=Q.FindField('T_LIBELLE').AsString;
      ARecord^.bGeneral:=False;
      result           :=ARecord^.Libelle ;
      General          :=false;
      ListeCompte.Add(ARecord);
    end ;
    Ferme (Q);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
function RechercheEcriture(L : TList;Date : TDateTime;Journal,Compte,CompteRef : string; Debit : boolean) : TLigneEcriture;
var ARecord : TLigneEcriture;
    i : integer;
begin
  Result := nil;
  for i := 1 to L.Count do
  begin
    ARecord := L.Items[i-1];
    if ((ARecord.Date=Date) and (ARecord.CodeJournal=Journal) and (ARecord.Compte=Compte)
       and ((CompteRef = '') or (ARecord.CompteRef=CompteRef)))
       and ( (Debit and ( Arecord.Debit<>0)) or ( not Debit and ( ARecord.Credit<>0))) then  // Pas débit et crédit renseigné sur la même ligne
    begin
      Result := ARecord;
      break;
    end;
  end;
end;


////////////////////////////////////////////////////////////////////////////////
procedure MajLigneEcriture (var L:TList;  Indice:integer;  CodeImmo:string; ParamEcr:TParamEcr;
                            Compte:string;  Montant:double;  Debit,bRappCpta:boolean;
                            NivDetail:integer;  AvecSurAmort:boolean);
var ARecord : TLigneEcriture; bGeneral: boolean;
    Montantecr : double;
    NouveauRecord, NouveauDebit : boolean;
    NouveauCompte : string;
begin
  ARecord := nil;

  // fq 16750  - qd on est en rapprochement il faut traiter l'enreg même quand le montant = 0 (cas des comptes associés)
  if (not bRappCpta) and (Montant=0.0) then exit;
  if (Compte='') then exit;

// BTY FQ 17748 Réajuster montant/compte/sens ICI avant recherche si écriture existe
// dans la Tlist car sinon ne serait pas regroupée si demandé
  NouveauDebit := Debit;
  NouveauCompte := Compte;
  NouveauRecord := False;

  IF fdotation = True Then
  begin
    // MVG 22/06/2006 FQ 10319
    if fprorata then Montantecr := RecalculMontant(Montant,ParamEcr.DateCalcul,ParamEcr.DateAmort,ParamEcr.DateFinAmort,ParamEcr.bDotation,ParamEcr.Methode, AvecSurAmort)
    else Montantecr := Montant;

    if Debit then
       Begin
       IF Montantecr<0 Then
          begin
          Montantecr := Abs(Montantecr);
          NouveauCompte := fcompte;
          NouveauDebit := not Debit;
          end
       Else
          begin
          Montantecr := Abs(Montantecr);
          NouveauDebit := Debit;
          end
       End
    else
       Begin
       IF Montantecr<0 Then
          begin
          Montantecr := Abs(Montantecr);
          NouveauDebit := not Debit;
          end
       Else
          begin
          Montantecr := Abs(Montantecr);
          NouveauDebit := Debit;
          end;
       End;
  end
  else
  // écriture autre que dotation
//   Montantecr := RecalculMontant(Montant,ParamEcr.DateCalcul,ParamEcr.DateAmort,ParamEcr.DateFinAmort,ParamEcr.bDotation,ParamEcr.Methode, AvecSurAmort);
   // MVG 22/06/2006 FQ 10319
   if fprorata then Montantecr := RecalculMontant(Montant,ParamEcr.DateCalcul,ParamEcr.DateAmort,ParamEcr.DateFinAmort,ParamEcr.bDotation,ParamEcr.Methode, AvecSurAmort)
   else Montantecr := Montant;

  case NivDetail of   // 0:synthétique, 1:par compte,  2:détaillé
    0:ARecord := RechercheEcriture(L,ParamEcr.Date,ParamEcr.Journal,
                 NouveauCompte,'',NouveauDebit);
    1:ARecord := RechercheEcriture(L,ParamEcr.Date,ParamEcr.Journal,
                 NouveauCompte,ParamEcr.CompteRef, NouveauDebit);
  end;

  if (ARecord = nil) then
//   //  Trame de l'écriture pas trouvée dans la Tlist => la créer

  begin
    ARecord := TLigneEcriture.Create;

    // BTY FQ 17748
    NouveauRecord := True;

    if ParamEcr.bDotation then
    case NivDetail of
      0: ARecord.Piece := '';
      1: ARecord.Piece := ParamEcr.CompteRef;
      //2: ARecord.Piece := CodeImmo;  FQ 15609
    end;
    if not ParamEcr.bDotation then ARecord.Piece := IntToStr(indice);
    // BTY FQ 15609 En niveau détaillé, référence facture de l'immo ou du crédit-bail
    if (NivDetail=2) then
       ARecord.Piece := ParamEcr.Ref_Immo;

//    CurPiece := ARecord.Piece;
    ARecord.CodeImmo := CodeImmo;
    ARecord.Credit := 0.0; ARecord.Debit := 0.0;
    ARecord.Date := ParamEcr.Date;
    ARecord.CodeJournal := ParamEcr.Journal;
    if NivDetail = 1 then ARecord.CompteRef := ParamEcr.CompteRef
    else ARecord.CompteRef := '';
    ARecord.Etablissement := ParamEcr.Etablissement;

    // BTY FQ 17748  + FQ 17760
    //ARecord.Libelle := RecupereLibelle(Compte,bGeneral);
    //if bGeneral then ARecord.Compte := Compte else ARecord.Auxi := Compte;
    ARecord.Libelle := RecupereLibelle(NouveauCompte, bGeneral);
    if bGeneral then ARecord.Compte := NouveauCompte
    else ARecord.Auxi := NouveauCompte;

    //  BTY FQ 15609 En niveau détaillé, libellé de l'immo
    if NivDetail = 2 then
       ARecord.Libelle := ParamEcr.Lib_Immo;
    if ParamEcr.Libelle <> '' then ARecord.Libelle := ParamEcr.Libelle;

// BTY FQ 17748 bloc déplacé plus haut
{    // TGA 16/11/2005
    // Réaffectation du montant de la dotation au débit ou crédit selon son signe
    IF fdotation = True Then
      Begin
        Montantecr := RecalculMontant(Montant,ParamEcr.DateCalcul,ParamEcr.DateAmort,ParamEcr.DateFinAmort,ParamEcr.bDotation,ParamEcr.Methode, AvecSurAmort);
        if Debit then
          Begin
           IF Montantecr<0 Then
             Begin
              ARecord.Credit:= Abs(Montantecr);
              ARecord.Compte := fcompte;
             End
           Else
              ARecord.Debit := Abs(Montantecr);
          End
        Else
          Begin
           IF Montantecr<0 Then
              ARecord.Debit := Abs(Montantecr)
           Else
              ARecord.Credit:= Abs(Montantecr);
          End;
      End
    // Fin TGA 16/11/2005
    Else
      Begin
        if Debit then ARecord.Debit :=RecalculMontant(Montant,ParamEcr.DateCalcul,ParamEcr.DateAmort,ParamEcr.DateFinAmort,ParamEcr.bDotation,ParamEcr.Methode, AvecSurAmort)
                 else ARecord.Credit:=RecalculMontant(Montant,ParamEcr.DateCalcul,ParamEcr.DateAmort,ParamEcr.DateFinAmort,ParamEcr.bDotation,ParamEcr.Methode, AvecSurAmort);
      End;
     }

  // BTY 03/06 FQ 17748
  end;

  if NouveauDebit then
     if NouveauRecord then  ARecord.Debit := MontantEcr
     else ARecord.Debit :=  ARecord.Debit + MontantEcr
  else
     if NouveauRecord then  ARecord.Credit := MontantEcr
     else ARecord.Credit := ARecord.Credit + MontantEcr;
  //

  {$IFDEF SERIE1}
  {$ELSE}
  // BTY 03/06 FQ 17713
  //if not bRappCpta then ChargeVentilImo (IsCompteVentilable(Compte),ARecord.Axes,'IM',CodeImmo,Montant,ARecord.Debit+ARecord.Credit);
  if not bRappCpta then
       ChargeVentilImo (IsCompteVentilable(NouveauCompte),ARecord.Axes,'IM',
                        CodeImmo, Abs(MontantEcr), ARecord.Debit+ARecord.Credit);
  {$ENDIF}


  if NouveauRecord then // BTY 03/06 FQ 17748
     L.Add(ARecord);

  // BTY 03/06 FQ 17748
  //end;
end;


////////////////////////////////////////////////////////////////////////////////
function SortiePendantPeriode (bOpeSortie : boolean; CodeImmo : string; DateCalcul : TDateTime; var ExcepSurSortie : double) : boolean;
var Q : TQuery;
begin
  result := false; ExcepSurSortie := 0;
  if not bOpeSortie then exit;
  Q := OpenSQL ('SELECT IL_IMMO,IL_DATEOP,IL_TYPEOP,IL_TYPEEXC,IL_MONTANTEXC FROM IMMOLOG WHERE IL_IMMO="'
        +CodeImmo+'" AND IL_DATEOP<="'+USDateTime(DateCalcul)+'" AND IL_DATEOP>="'+USDateTime(VHImmo^.Encours.Deb)+'" '
        +'AND IL_TYPEOP LIKE "CE%" ORDER BY IL_DATEOP DESC',True);
  while not Q.Eof do
  begin
    if Q.FindField('IL_TYPEEXC').AsString <> '' then // mbo 25.11.05 FQ 17084 - test sur le type
    begin
      if Q.FindField('IL_TYPEEXC').AsString = 'DOT' then
        ExcepSurSortie := ExcepSurSortie+ Q.FindField('IL_MONTANTEXC').AsFloat
      else
        ExcepSurSortie := ExcepSurSortie - Q.FindField('IL_MONTANTEXC').AsFloat;
    end;
    result := True;
    Q.Next;
  end;
  Ferme (Q);
end;

////////////////////////////////////////////////////////////////////////////////
function MutationPendantPeriode (Q: TQuery;bOpeMutation : boolean; CodeImmo : string; DateCalcul : TDateTime; var mtSortie : double; var CpteMutation : string) : boolean;
var QOPe : TQuery;
    Plan : TPlanAmort;
begin
  result := false; mtSortie := 0; CpteMutation := '';
  if not bOpeMutation then exit;
  QOpe := OpenSQL ('SELECT IL_IMMO,IL_CPTEMUTATION,IL_DATEOP,IL_PLANACTIFAP,IL_TYPEOP FROM IMMOLOG WHERE IL_IMMO="'+
        CodeImmo+'" AND IL_DATEOP<="'+USDateTime(DateCalcul)+'" AND IL_DATEOP>="'+
        USDateTime(VHImmo^.Encours.Deb)+'" ORDER BY IL_DATEOP DESC',True);
  if not QOpe.Eof then
  begin
    while ((not QOpe.Eof) and (QOpe.FindField('IL_CPTEMUTATION').AsString = '')) do QOpe.Next;
    if not QOpe.Eof then
    begin
      CpteMutation := QOpe.FindField ('IL_CPTEMUTATION').AsString;
      while ((not QOpe.Eof) and (QOpe.FindField('IL_TYPEOP').AsString<>'CES') and (QOpe.FindField('IL_TYPEOP').AsString<>'CEP')) do QOpe.Next;
      if not QOpe.Eof then
      begin
        Plan:=TPlanAmort.Create(true) ;// := CreePlan(True);
        try
          Plan.Charge (Q);
          Plan.Recupere (CodeImmo,QOpe.FindField ('IL_PLANACTIFAP').AsString);
          mtSortie := Plan.GetCessionExercice(DateCalcul,Plan.AmortEco);
        finally
          Plan.free ; //Detruit;
        end ;
      end else mtSortie := 0;
      result := True;
    end;
  end else mtSortie := 0;
  Ferme (QOpe);
end;

////////////////////////////////////////////////////////////////////////////////
// BTY 10/05 En CRC 2002-10, test opération de modif base dans l'exercice
function ModifBasePendantPeriode (Q : TQuery ; bOpeModifBase : boolean; CodeImmo : string; DateCalcul : TDateTime; var DeltaReprise : double; bSortie:boolean) : boolean;
var QLog : TQuery;
begin
  Result := false; DeltaReprise := 0;
  if not bOpeModifBase then exit;

  QLog := OpenSQL ('SELECT IL_IMMO,IL_DATEOP,IL_TYPEOP,IL_MONTANTDOT,IL_MONTANTECL FROM IMMOLOG WHERE IL_IMMO="'  + CodeImmo
               + '" AND IL_DATEOP<="'+USDateTime(DateCalcul)+'" AND IL_DATEOP>="'+USDateTime(VHImmo^.Encours.Deb)+'" '
               + 'AND IL_TYPEOP="MB2" ORDER BY IL_DATEOP DESC',True);
  if not QLog.Eof then
     begin
     if bSortie then
        DeltaReprise := Q.FindField('I_REPCEDFISC').AsFloat - Q.FindField('I_REPCEDECO').AsFloat
     else
        DeltaReprise := Q.FindField('I_REPRISEFISCAL').AsFloat - Q.FindField('I_REPRISEECO').AsFloat;

     // BTY 12/05 Attention OK uniquement si une seule modif de bases dans l'exercice
     // car ici on n'est positionné que sur l'enreg IMMOLOG le plus récent
     // et plusieurs modifs bases autorisées par exercice
     DeltaReprise := DeltaReprise - (QLog.FindField ('IL_MONTANTECL').AsFloat - QLog.FindField ('IL_MONTANTDOT').AsFloat);

     end;
  Result := (not QLog.Eof);
  Ferme (QLog);
end;


// BTY 11/05 Test opération de révision du plan CD2 (ajout plan fiscal CRC 2002-10) dans l'exercice
function RevPlanCRCPendantPeriode (Q : TQuery ; bOpeChangePlan : boolean; CodeImmo : string; DateCalcul : TDateTime; var DeltaPlanCRC : double) : boolean;
var QLog : TQuery;
begin
  Result := false; DeltaPlanCRC := 0;
  if not bOpeChangePlan then exit;

  QLog := OpenSQL ('SELECT IL_IMMO,IL_DATEOP,IL_TYPEOP FROM IMMOLOG WHERE IL_IMMO="'  + CodeImmo
               + '" AND IL_DATEOP<="'+USDateTime(DateCalcul)+'" AND IL_DATEOP>="'+USDateTime(VHImmo^.Encours.Deb)+'" '
               + 'AND IL_TYPEOP="CD2" ORDER BY IL_DATEOP DESC',True);
  if  not QLog.Eof then DeltaPlanCRC := Q.FindField('I_REPRISEFISCAL').AsFloat - Q.FindField('I_REPRISEECO').AsFloat;
  Result := (not QLog.Eof);
  Ferme (QLog);
end;

// PGR 12/2005 Opération de Changement de méthode CDA dans l'exercice
Function ChgtMethodePendantPeriode ( Plan : TPlanAmort; bOpeChangePlan : boolean; CodeImmo : string; DateCalcul : TDateTime; bSortie:boolean; var DeltaEco : double; var DeltaFisc : double) : boolean;
var QLog : TQuery;
    NewEco : double;
    NewFisc : double;

begin
  Result := false; DeltaEco := 0; DeltaFisc := 0;

  if not bOpeChangePlan then exit;

  QLog := OpenSQL ('SELECT IL_IMMO,IL_DATEOP,IL_TYPEOP,IL_REPRISEECO,IL_REPRISEFISC,IL_REVISIONREPECO,IL_REVISIONREPFISC FROM IMMOLOG WHERE IL_IMMO="'  + CodeImmo
               + '" AND IL_DATEOP<="'+USDateTime(DateCalcul)+'" AND IL_DATEOP>="'+USDateTime(VHImmo^.Encours.Deb)+'" '
               + 'AND IL_TYPEOP="CDA" ORDER BY IL_DATEOP DESC',True);
  if  not QLog.Eof then
  begin

    if bsortie then
    begin
      NewEco := Plan.AmortEco.RepriseCedee;
      NewFisc := Plan.AmortFisc.RepriseCedee;
    end
    else
    begin
      NewEco := Plan.AmortEco.Reprise;
      NewFisc := Plan.AmortFisc.Reprise;
    end;

    DeltaEco := NewEco - QLog.FindField('IL_REPRISEECO').AsFloat;

    if (QLog.FindField('IL_REVISIONREPECO').AsFloat <> 0) OR (NewFisc <> 0) then
      DeltaFisc := (NewEco - NewFisc)
               - (QLog.FindField('IL_REPRISEECO').AsFloat - QLog.FindField('IL_REVISIONREPECO').AsFloat);
  end;

  Result := (not QLog.Eof);
  Ferme (QLog);
end;

// PGR - 12/2005 - Opération de Changement de méthode (Eclatement préalable) ACQ dans l'exercice
Function ChgtMethodeEclatPrealable (Q : TQuery; bFille : boolean; DateCalcul : TDateTime; bSortie:boolean; var DeltaHt : double) : boolean;
var QLog : TQuery;
    CodeImmo : string;

begin
  Result := false; DeltaHt := 0;

  if not bFille then exit;

  CodeImmo := Q.FindField('I_IMMO').AsString;
  QLog := OpenSQL ('SELECT IL_IMMO,IL_DATEOPREELLE,IL_TYPEOP,IL_MONTANTECL FROM IMMOLOG WHERE IL_IMMO="'  + CodeImmo
               + '" AND IL_DATEOPREELLE<="'+USDateTime(DateCalcul)+'" AND IL_DATEOPREELLE>="'+USDateTime(VHImmo^.Encours.Deb)+'" '
               + 'AND IL_TYPEOP="ACQ" ORDER BY IL_DATEOP DESC',True);

  if (not QLog.Eof) then
  begin
    if Qlog.FindField('IL_MONTANTECL').AsFloat <> 0 then
      DeltaHt := Q.FindField('I_VALEURACHAT').AsFloat - QLog.FindField('IL_MONTANTECL').AsFloat;
  end;

  if DeltaHt <> 0 then
    Result := true;
  Ferme (QLog);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 02/11/2006
Modifié le ... :   /  /
Description .. : On récupère l'ancien montant de la subvention dans
Suite ........ : l'immolog correspondant à l'enregistrement de l'opération de
Suite ........ : réduction de la subvention d'investissement
Mots clefs ... :
*****************************************************************}
function ReducSBVPendantPeriode (Q : TQuery ; bOperation : boolean; CodeImmo : string; DateCalcul : TDateTime; var OldMntSBV : double; bSortie:boolean) : boolean;
var QLog : TQuery;
begin
  Result := false;
  OldMntSBV := 0;
  if not bOperation then exit;

  QLog := OpenSQL ('SELECT IL_IMMO,IL_DATEOP,IL_MONTANTEXC FROM IMMOLOG WHERE IL_IMMO="'  + CodeImmo
               + '" AND IL_DATEOP<="'+USDateTime(DateCalcul)+'" AND IL_DATEOP>="'+USDateTime(VHImmo^.Encours.Deb)+'" '
               + 'AND IL_TYPEOP="RSB"',True);
  if not QLog.Eof then
     OldMntSBV := QLog.FindField('IL_MONTANTEXC').AsFloat;

  Result := (not QLog.Eof);
  Ferme (QLog);
end;


////////////////////////////////////////////////////////////////////////////////
function GetVNCNam (Code : string; DateCalcul : TDateTime) : double;
var Q : TQuery;
    stSelect, stWhere : string;
    VNC : double;
begin
  VNC := 0;
  stSelect := 'SELECT IL_IMMO,IL_DATEOP,IL_TYPEOP,IL_VOCEDEE FROM IMMOLOG';
  stWhere := ' WHERE IL_IMMO="'+Code+'" AND IL_DATEOP<="'+USDateTime(DateCalcul)+
        '" AND IL_DATEOP>="'+USDateTime(VHImmo^.Encours.Deb)+'" AND (IL_TYPEOP="CES" OR IL_TYPEOP="CEP")';
  Q := OpenSQL (stSelect + stWhere, True );
  while not Q.Eof do
  begin
    VNC := VNC + Q.FindField ('IL_VOCEDEE').AsFloat;
    Q.Next;
  end;
  Ferme (Q);
  Result := VNC;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 01/06/2004
Modifié le ... :   /  /
Description .. : - LG - 01/06/2004 - Modif de l'appel de la fct
Suite ........ : RecupInfoTVA
Mots clefs ... :
*****************************************************************}
procedure  CalculIntegrationEcheance (Q : TQuery; var L : TList; ParamEch,ParamPai : TParamEcr;NivDetail : integer; dtEchMin,dtEchMax : TDateTime;AvecDejaIntegre: boolean);
var QEcheance : TQuery ;
    TTC, Montant : double;
    ImmoEnCours,LeSql : string;
    Indice : integer;
    {$IFDEF SERIE1}
    {$ELSE}
    StCompteTva : string;
    stGene, stAuxi : string;
    LesComptes : TZCompte;
    LesTiers : TZTiers;
    stMemGestionTVA : string;
    LCpteTvaOk : HTStrings;
    LCpteTvaPasOk : HTStrings;
    {$ENDIF}
    RdTauxTva : double;

begin
  Indice := 0;
  {$IFDEF SERIE1}
  {$ELSE}
  LesComptes := TZCompte.Create;
  LesTiers := TZTiers.Create;
  LCpteTvaOk := HTStringList.Create;
  LCpteTvaPasOk := HTStringList.Create;
  {$ENDIF}
  try
    ImmoEnCours :=Q.FindField('I_IMMO').AsString;

    // BTY FQ 15609 Stocker libellé immo et référence facture
    if (ParamEch<>nil) then
        begin
        ParamEch.Lib_Immo   :=Q.FindField ('I_LIBELLE').AsString;
        ParamEch.Ref_Immo   :=Q.FindField ('I_REFINTERNEA').AsString;
        end;
    LeSql:='SELECT * FROM IMMOECHE WHERE IH_IMMO="'+ImmoEnCours +'" '
          +' AND IH_DATE <="'+USDateTime(dtEchMax)+'" AND IH_DATE >="'+USDateTime(dtEchMin)+'" ' ;
    if not AvecDejaIntegre then LeSql:=LeSql+' AND (IH_INTEGREECH="-" OR IH_INTEGREPAI="-")' ;
    QEcheance:=OpenSQL (LeSql,TRUE);
    while not QEcheance.Eof do
    begin
      Montant := QEcheance.FindField('IH_MONTANT').AsFloat;
      Montant := AffecteTVAEch(Q,Q.FindField ('G_TVA').AsString,Montant);
      if (ParamEch<>nil) and (ParamEch.Journal <> '')
         and ((QEcheance.FindField('IH_INTEGREECH').AsString='-') or AvecDejaIntegre) then
      begin
        Inc(Indice);
        ParamEch.Etablissement := Q.FindField ('I_ETABLISSEMENT').AsString;
        ParamEch.Date := QEcheance.FindField('IH_DATE').AsDateTime;
        {$IFDEF SERIE1}
        RdTauxTva:=ImCompte2Taux(ParamEch.CompteRef,true) ;
        {$ELSE}
        stMemGestionTVA := ParamEch.CompteRef;
        if ParamEch.CompteRef = 'X' then // Gestion de la TVA
        begin
          stGene := Q.FindField('I_COMPTELIE').AsString;
          stAuxi := Q.FindField('I_ORGANISMECB').AsString;
          LesComptes.GetCompte( stGene );
          LesTiers.GetCompte( stAuxi );
          LesComptes.RecupInfoTVA(Q.FindField('I_COMPTELIE').AsString,
                                  Montant,
                                  '',
                                  '',
                                  '',
                                  StCompteTva,
                                  RdTauxTva,
                                  LesTiers.GetValue('T_REGIMETVA',0));
          ParamEch.CompteRef := BourreEtLess(StCompteTva, fbGene);
          if (LCpteTvaOk.IndexOf(ParamEch.CompteRef) < 0) then
          begin
            if (LCpteTvaPasOk.IndexOf(ParamEch.CompteRef) < 0) then
            begin
              if ExisteSQL ('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL="'+ParamEch.CompteRef+'"') then
                LCpteTvaOk.Add(ParamEch.CompteRef)
              else begin  // Compte de TVA n'existe pas
                LCpteTvaPasOk.Add(ParamEch.CompteRef);
                ParamEch.CompteRef := '';
              end;
            end else ParamEch.CompteRef := ''; // Compte de TVA n'existe pas
          end;

        end else RdTauxTva := 0;
        {$ENDIF}
        // FQ 14338 - CA - 26/08/2004 HT:=Montant/(1.0+(RdTauxTva/100.0)) ;
        TTC:=Montant*(1.0+(RdTauxTva/100.0)) ;
        MajLigneEcriture(L,Indice,ImmoEnCours,ParamEch,Q.FindField('I_COMPTELIE').AsString,Montant,true,false,NivDetail,False);
        MajLigneEcriture(L,Indice,ImmoEnCours,ParamEch,ParamEch.CompteRef,TTC-Montant,true,false,NivDetail,False);
        MajLigneEcriture(L,Indice,ImmoEnCours,ParamEch,Q.FindField('I_ORGANISMECB').AsString,TTC,false,false,NivDetail, False);
        {$IFDEF SERIE1}
        {$ELSE}
        ParamEch.CompteRef := stMemGestionTVA;
        {$ENDIF}
      end;
      if (ParamPai<>nil) and (ParamPai.Journal<>'')
          and ((QEcheance.FindField('IH_INTEGREPAI').AsString='-') or AvecDejaIntegre) then
      begin
        if ParamPai.Journal<>ParamEch.Journal then Inc(Indice);
        ParamPai.Etablissement:=Q.FindField ('I_ETABLISSEMENT').AsString;
        ParamPai.Date:=QEcheance.FindField('IH_DATE').AsDateTime;
        MajLigneEcriture(L,Indice,ImmoEnCours,ParamPai,Q.FindField('I_ORGANISMECB').AsString,Montant,true,false,NivDetail, False);
        MajLigneEcriture(L,Indice,ImmoEnCours,ParamPai,ParamPai.CompteRef,Montant,false,false,NivDetail, False);
      end;
      QEcheance.Next;
    end;
    Ferme(QEcheance);
  finally
    {$IFDEF SERIE1}
    {$ELSE}
    LCpteTvaOk.Free;
    LCpteTvaPasOk.Free;
    LesTiers.Free;
    LesComptes.Free;
    {$ENDIF}
  end;
end;

////////////////////////////////////////////////////////////////////////////////
procedure  CalculIntegrationDotation ( Q : TQuery; var L : TList; ParamEcr : TParamEcr;NivDetail : integer);
var CAss : TCompteAss;
    dDotCalc,dDotEco,dDot,dCesEco,dExcEco,mDerog,mDerogCes,CumulCesEco,CumulCesFisc: double;
    CumulEco,CumulFisc, CumulDerog : double;
    VNC :double;
    ImmoEnCours, CpteMutation : string;
    Indice : integer;
    Plan : TPlanAmort;
    bSortie,bCessionTotale, bMutation : boolean;
    PlanValeur,CessionAvantMutation,ExcepSurSortie : double;
    bRappCpta,bNam : boolean ;
    DateOpe : TDateTime;
    PlanInfo : TPlanInfo;
    bModifBase : boolean;
    DeltaReprise : double;
    bRevPlanCRC : boolean;
    DeltaPlanCRC : double;
    bChgtMethode : boolean;
    DeltaEco : double;
    DeltaFisc : double;
    bEclatPrealable : boolean;
    DeltaHt : double;
    VA : double;
    //ajout mbo pour la subvention
    CSbv : TCompteSBV;
    bReducSBV : boolean;
    OldMntSBV : double;
    DeltaReduc : double;
    Sbvsortie : double;
    bfille : boolean;
begin

  bReducSBV := false;
  DeltaReduc:= 0.00;
  // TGA 16/11/2005
  fdotation := false;
  fprorata := true; // MVG 22/06/2006

  Dotation_V2 := False;
  bRappCpta:=(NivDetail=4);
  //TEST YCP 22/10/01 if (Q.FindField('I_DATEAMORT').AsDateTime>ParamEcr.DateCalcul) then exit; // Pas d'écritures si amort. pas commencé
  if (Q.FindField('I_DATEAMORT').AsDateTime>ParamEcr.DateCalcul)
  and (Q.FindField('I_DATEPIECEA').AsDateTime>ParamEcr.DateCalcul) then exit;

  ImmoEnCours   :=Q.FindField('I_IMMO').AsString;
  bCessionTotale:=((Q.FindField ('I_QUANTITE').AsInteger = 0) and (Q.FindField ('I_OPECESSION').AsString = 'X'));
  bMutation     :=MutationPendantPeriode (Q,Q.FindField ('I_OPEMUTATION').AsString ='X',ImmoEnCours,ParamEcr.DateCalcul,CessionAvantMutation,CpteMutation);
  bSortie       :=SortiePendantPeriode (Q.FindField ('I_OPECESSION').AsString ='X',ImmoEnCours,ParamEcr.DateCalcul,ExcepSurSortie);
  Indice        :=0;
  bNam          :=(Q.FindField('I_METHODEECO').AsString='NAM');

  // BTY 10/05
  bModifBase    := ModifBasePendantPeriode (Q,Q.FindField ('I_OPEMODIFBASES').AsString ='X',ImmoEnCours,ParamEcr.DateCalcul, DeltaReprise, bSortie);
  bRevPlanCRC   := RevPlanCRCPendantPeriode(Q,Q.FindField ('I_OPECHANGEPLAN').AsString ='X',ImmoEnCours,ParamEcr.DateCalcul, DeltaPlanCRC);

  //Creation du plan
  Plan:=TPlanAmort.Create(true) ;
  Plan.Charge (Q);
  // Récupére pour les dates strictement inférieures à la date passée en paramètre. On ajoute donc 1 à la date
  // pour prendre en compte les opérations inférieures qui ont eu lieu à une date inférieure ou
  // égale à la date passée en paramètre.
  Plan.RecupereSuivantDate(ParamEcr.DateCalcul+1);
  Plan.CalculDateFinAmortissement(Plan.AmortEco);
  ParamEcr.DateFinAmort:=Plan.AmortEco.DateFinAmort;
  ParamEcr.DateAmort   :=Q.FindField ('I_DATEAMORT').AsDateTime;
  RecupereComptesAssocies (Q,'',CAss);
  ParamEcr.CompteRef    :=CAss.Immo;

  // fq 16750 ajout mbo on crée à vide tous les enreg comptes associés
  if bRappCpta then
  begin
   if Cass.Immo <> '' then MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.Immo,0,true,bRappCpta,0,False);
   if Cass.Amort <> '' then MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.Amort,0,true,bRappCpta,0,False);
   if Cass.Dotation <> '' then MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.Dotation,0,true,bRappCpta,0,False);
   if Cass.Derog <> '' then MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.Derog,0,true,bRappCpta,0,False);
   if Cass.RepriseDerog <> '' then MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.RepriseDerog,0,true,bRappCpta,0,False);
   if Cass.DotationExcep <> '' then MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.DotationExcep,0,true,bRappCpta,0,False);
   if Cass.ProvisDerog <> '' then MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.ProvisDerog,0,true,bRappCpta,0,False);
   if Cass.VaCedee <> '' then MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.VaCedee,0,true,bRappCpta,0,False);
   if Cass.AmortCede <> '' then MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.AmortCede,0,true,bRappCpta,0,False);
   if Cass.VoaCede <> '' then MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.VoaCede,0,true,bRappCpta,0,False);
   if Cass.RepriseExcep <> '' then MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.RepriseExcep,0,true,bRappCpta,0,False);
   if Cass.RepriseExploit <> '' then MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.RepriseExploit,0,true,bRappCpta,0,False);
  end;
  // fin ajout

  ParamEcr.Etablissement:=Q.FindField ('I_ETABLISSEMENT').AsString;
  dDotEco :=Plan.GetDotationGlobale(1,VHImmo^.Encours.Fin,dDot,dCesEco,dExcEco);
  // BTY FQ 15609 Stocker libellé de l'immo
  ParamEcr.Lib_Immo   :=Q.FindField ('I_LIBELLE').AsString;
  // BTY FQ 15609 Stocker référence facture immo
  ParamEcr.Ref_Immo   :=Q.FindField ('I_REFINTERNEA').AsString;

  // MVG 22/06/06 FQ 10319
  VA := Q.FindField ('I_VALEURACHAT').AsFloat;

  //mbo - 24.10.05 pour crc 2002-10 dExcEco :=Plan.GetExcepExercice(ParamEcr.DateCalcul, DateOpe) ;
  /////dExcEco :=Plan.GetExcepExercice(ParamEcr.DateCalcul, DateOpe, false,true) ;

  //PGR 12/2005 Changement de méthode(Changement des conditions d'amortissement)
  bChgtMethode  := ChgtMethodePendantPeriode(Plan,Q.FindField ('I_OPERATION').AsString ='X',ImmoEnCours,ParamEcr.DateCalcul, bSortie, DeltaEco, DeltaFisc);

  //PGR 12/2005 Changement de méthode(Eclatement préalable)
  bfille:=Q.FindField ('I_IMMOORIGINEECL').AsString <> '';
  if (bfille=true) then // Est-ce un éclatement préalable au chgt de méthode FQ19956 et 19969
    begin
    if not ExisteSQL ('SELECT IL_TYPEOP FROM IMMOLOG WHERE IL_IMMO="'+ImmoEnCours+'"'+' AND IL_TYPEOP="CDA"') then bfille:=false;
    end
  else
    begin
      bfille:=false;
    end;

  bEclatPrealable  := ChgtMethodeEclatPrealable(Q,bfille,ParamEcr.DateCalcul, bSortie, DeltaHt);

  //mbo - 02.11.06 - ajout pour subvention d'équipement
  if Plan.SBV then
  begin
     GetComptesSBV(CSbv, ImmoEnCours);  // charge les comptes de subvention
     if bRappCpta then
     begin
        // on crée à vide les enreg compte de subvention
        if CSbv.Dotation <> '' then MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CSbv.Dotation,0,true,bRappCpta,0,False);
        if CSbv.Reprise <> '' then MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CSbv.Reprise,0,true,bRappCpta,0,False);
        if CSbv.QuotePart <> '' then MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CSbv.QuotePart,0,true,bRappCpta,0,False);
     end;
     bReducSBV := ReducSBVPendantPeriode(Q,Q.FindField ('I_OPERATION').AsString ='X',ImmoEnCours,ParamEcr.DateCalcul, OldMntSBV, bSortie);
     // sbvsortie:=Q.FindField ('I_SBVMTC').Asfloat; MVG 10/01/07 Conseil compil
     If (bsortie) then
     begin
         sbvsortie:=Q.FindField ('I_SBVMTC').Asfloat;
         DeltaReduc := OldMntSBV - sbvsortie;
     end
     else
         begin
         DeltaReduc := OldMntSBV - Plan.MntSBV;
     end;
  end;

  // TGA 07/11/2005 pas de dépréciation dans l'exceptionnel
  dExcEco :=Plan.GetExcepExercice(ParamEcr.DateCalcul, DateOpe, false,false) ;

  Plan.GetCumulsDotExercice(VHImmo^.Encours.Deb,CumulEco,CumulFisc, false, True, false);

  if bRappCpta then
    begin
    // On génère la ligne correspondant au montant HT en se ramenant au cas général synthétique
    fprorata:=false; //MVG 22/06/2006 FQ 10319
    ParamEcr.Methode:='LIN';
    //MVG 22/06/2006 FQ 10319
    if (not bcessiontotale) then
       begin
       PlanValeur:=Plan.ValeurHT+Plan.ValeurTVARecuperable-Plan.ValeurTvaRecuperee ; //AmortEco.Base+Plan.GetBaseCessionExercice(VHImmo^.Encours.Fin,Plan.AmortEco);
       MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.Immo,PlanValeur,true,bRappCpta,0,False);
       end
    else
      begin
         if (not bsortie) then MajLigneEcriture(L,Indice,ImmoEnCours,ParamEcr,CAss.VoaCede  ,VA,true,bRappCpta,NivDetail, False);
      end;
    NivDetail := 0;
    fprorata:=true; //MVG 22/06/2006 FQ 10319
    end;

  //  if (Q.FindField('I_DATEAMORT').AsDateTime>ParamEcr.DateCalcul) then exit;
  // CA - 29/04/2002 - Prise en compte fiscal par rapport à date de mise en service

  //mbo - 02.11.06 - ce test est déjà fait en début de fonction
  //if (Q.FindField('I_DATEAMORT').AsDateTime>ParamEcr.DateCalcul)
  //and (Q.FindField('I_DATEPIECEA').AsDateTime>ParamEcr.DateCalcul) then exit;

  ParamEcr.Methode := Plan.AmortEco.Methode;

  // --- Mutation --------------------------------------------------------------
  if bMutation then
    begin
    ParamEcr.Methode:='NUL';
    if not bRappCpta then
      begin
      if bNam then
        begin
        MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CpteMutation, Q.FindField('I_MONTANTHT').AsFloat,false,bRappCpta,NivDetail,False);
        MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.Immo,Q.FindField('I_MONTANTHT').AsFloat,true,bRappCpta,NivDetail, False);
        end else
        begin
        MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CpteMutation, Plan.BaseDebutExoEco,false,bRappCpta,NivDetail, Plan.SurAmort);
        MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.Immo,Plan.BaseDebutExoEco,true,bRappCpta,NivDetail, Plan.SurAmort);
        MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAssAmortissement(CpteMutation), CumulEco,true,bRappCpta,NivDetail, Plan.SurAmort);
        MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.Amort,CumulEco,false,bRappCpta,NivDetail, Plan.SurAmort);
        end;
     end;
    ParamEcr.Methode := Plan.AmortEco.Methode;
    end;

  // --- Dotation --------------------------------------------------------------
  // Tga 16/11/2005 pour test si mt négatif
  fdotation := True;
  fcompte := CAss.RepriseExploit;

  if bSortie and bCessionTotale then
   begin
    ParamEcr.Methode := 'NUL';
    if bRappCpta then
    begin
      MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.Dotation, dDotEco-(dExcEco+ExcepSurSortie),true,bRappCpta,NivDetail, False);
      MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.Amort,0,false,bRappCpta,NivDetail, False);
      //MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.Amort,dDotEco+CumulEco+dExcEco+ExcepSurSortie,false,NivDetail);
    end else
    begin
      MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.Dotation, dDotEco-(dExcEco+ExcepSurSortie),true,bRappCpta,NivDetail,False);
      MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.Amort,dDotEco(*+dExcEco+ExcepSurSortie*),false,bRappCpta,NivDetail,False);
    end;
    ParamEcr.Methode := Plan.AmortEco.Methode;
   end
  else
   begin
    dDotCalc := RecalculMontant (dDotEco-dExcEco-dCesEco,ParamEcr.DateCalcul,ParamEcr.DateAmort,ParamEcr.DateFinAmort,TRUE,Plan.AmortEco.Methode,Plan.Suramort);
    ParamEcr.Methode := 'NUL';
    if bRappCpta then
    begin
      // compte dotation classe 6 = dotation de l'exercice
      MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.Dotation, dDotCalc+dCesEco-ExcepSurSortie,true,bRappCpta,NivDetail,False);

      // compte de classe 28 = antérieurs + dotation de l'exo   FQ 177773 -
      MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.Amort,dDotCalc+CumulEco+dExcEco{+dCesEco}-CessionAvantMutation,false,bRappCpta,NivDetail,False);
      //MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.Amort,dDotCalc+dExcEco{+dCesEco}-CessionAvantMutation,false,bRappCpta,NivDetail,False);

    end else
    begin
      MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.Dotation, dDotCalc+dCesEco-ExcepSurSortie,true,bRappCpta,NivDetail, False);
      MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.Amort,dDotCalc+dExcEco+dCesEco-CessionAvantMutation,false,bRappCpta,NivDetail,False);
    end;

    MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAssAmortissement(CpteMutation),CessionAvantMutation,false,bRappCpta,NivDetail,False);
    ParamEcr.Methode := Plan.AmortEco.Methode;

   end;
  fdotation := false;

  // --- Exceptionnel ----------------------------------------------------------
  ParamEcr.Methode := 'NUL';
  if (dExcEco+ExcepSurSortie) > 0 then
    MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.DotationExcep,dExcEco+ExcepSurSortie,true,bRappCpta,NivDetail, False)
  else
    MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.RepriseExcep,(-1)*(dExcEco+ExcepSurSortie),false,bRappCpta,NivDetail, False);
  ParamEcr.Methode:=Plan.AmortEco.Methode;
  mDerogCes := 0;
  // --- Dérogatoire -----------------------------------------------------------
  if Plan.Fiscal then
  begin
    PlanInfo := TPlanInfo.Create('');
    try
      PlanInfo.ChargeImmo(ImmoEnCours);
      PlanInfo.Calcul(ParamEcr.DateCalcul, true, false, ''); // fq 21754
      mDerog := PlanInfo.Derogatoire;
      if bSortie then mDerogCes := mDerog+PlanInfo.GetCumulAntDerogatoire(ParamEcr.DateCalcul);
      ParamEcr.Methode := 'NUL';
      if mDerog > 0 then MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.Derog,mDerog,false,bRappCpta,NivDetail,False)
      else if mDerog <0 then MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.Derog,(-1)*mDerog,true,bRappCpta,NivDetail,False);
      if mDerog > 0 then MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.ProvisDerog,mDerog,true,bRappCpta,NivDetail,False)
      else if mDerog <0 then MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.RepriseDerog,(-1)*mDerog,false,bRappCpta,NivDetail, False);
      if bRappCpta then
      begin
        CumulDerog := PlanInfo.GetCumulAntDerogatoire(ParamEcr.DateCalcul);
        MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.Derog,CumulDerog,false,bRappCpta,NivDetail, False);
      end;
      ParamEcr.Methode := Plan.AmortEco.Methode;
    finally
      PlanInfo.Free;
    end;
  end;



  // --- Sortie ----------------------------------------------------------------
  if bSortie then
  begin
    ParamEcr.Methode := 'NUL';
    Plan.GetCumulDotationCession(CumulCesEco, CumulCesFisc,VHImmo^.Encours.fin);

    // mbo - 25/08/06 - conseil compil - dExcAntEco:=Plan.GetExcepAnter(VHImmo^.Encours.deb);
    // mbo le 23.06.06 - fq 18470 - CumulCesEco :=CumulCesEco +Plan.AmortEco.RepriseCedee+dExcAntEco;

    CumulCesEco :=CumulCesEco +Plan.AmortEco.RepriseCedee;
    CumulCesFisc:=CumulCesFisc+Plan.AmortFisc.RepriseCedee;
    if bNam then VNC:=GetVNCNam(ImmoEnCours, ParamEcr.DateCalcul)
            else VNC:=Plan.GetBaseCessionExercice (VHImmo^.Encours.Fin,Plan.AmortEco)-CumulCesEco ;
//    if CumulCesFisc <> 0 then mDerogCes:=CumulCesFisc-CumulCesEco else mDerogCes:=0;
    if bCessionTotale then
      begin
      if VNC<0 then VNC:=0 ; //YCP 16-07-01
      //mbo 25.11.05 fQ 16400 if VNC<>0 then MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.VaCedee,VNC-dExcEco,True,bRappCpta,NivDetail, False);
      if VNC<>0 then MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.VaCedee,VNC,True,bRappCpta,NivDetail, False);
      if not bRappCpta then
        begin
        // mbo 25.11.05 MajLigneEcriture(L,Indice,ImmoEnCours,ParamEcr,CAss.AmortCede,CumulCesEco+dExcEco,true ,bRappCpta,NivDetail, False);
        MajLigneEcriture(L,Indice,ImmoEnCours,ParamEcr,CAss.AmortCede,CumulCesEco,true ,bRappCpta,NivDetail, False);
        MajLigneEcriture(L,Indice,ImmoEnCours,ParamEcr,CAss.VoaCede  ,VNC+CumulCesEco    ,false,bRappCpta,NivDetail, False);
        end
      end else
      begin
      // !! Calculer à partir de IL_VOCEDE pour ne pas avoir de pb d'arrondi : 29/01/2001
        MajLigneEcriture(L,Indice,ImmoEnCours,ParamEcr,CAss.VaCedee  ,abs(VNC)          ,(VNC>=0) ,bRappCpta,NivDetail, False);

        if not bRappCpta then MajLigneEcriture(L,Indice,ImmoEnCours,ParamEcr,CAss.AmortCede,CumulCesEco    ,true ,bRappCpta,NivDetail, False);
        MajLigneEcriture(L,Indice,ImmoEnCours,ParamEcr,CAss.VoaCede  ,VNC+CumulCesEco,false,bRappCpta,NivDetail, False);
      end;

    if mDerogCes > 0 then
    begin
      (*if not bRappCpta then *)MajLigneEcriture(L,Indice,ImmoEnCours,ParamEcr,CAss.Derog,mDerogCes,true ,bRappCpta,NivDetail, False);
      MajLigneEcriture(L,Indice,ImmoEnCours,ParamEcr,CAss.RepriseDerog,mDerogCes,false,bRappCpta,NivDetail, False);
    end
    else if mDerogCes < 0 then
    begin
      (* if not bRappCpta then *)MajLigneEcriture(L,Indice,ImmoEnCours,ParamEcr,CAss.Derog,(-1)*mDerogCes,false,bRappCpta,NivDetail, False);
      MajLigneEcriture(L,Indice,ImmoEnCours,ParamEcr,CAss.ProvisDerog,(-1)*mDerogCes,true,bRappCpta,NivDetail, False);
    end;
    // Ecritures TVA
    {15/12/00
    PlanInfo := TPlanInfo.Create(ImmoEnCours);
    PlanInfo.Plan.Copie(Plan);
    PlanInfo.Calcul(ParamEcr.DateCalcul,True);
    }
    (* YCP Retirer le 13/03/01
    if PlanInfo.TVAAReverser <> 0 then
    begin
      //EPZ 28/11/00
      { A FAIRE
        MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.CreanceCessionActif,PlanInfo.TotalPrixVente,True,NivDetail);
        MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.ProduitCessionActif,PlanInfo.TotalPrixVente-PlanInfo.TVAAReverser ,false,NivDetail);
        MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.TVACollectee,PlanInfo.TVAAReverser,false,NivDetail);
      }
      MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,ImBourreLaDoncSurLesComptes('462'),PlanInfo.TotalPrixVente,True,NivDetail);
      MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,ImBourreLaDoncSurLesComptes('775'),PlanInfo.TotalPrixVente-PlanInfo.TVAAReverser ,false,NivDetail);
      MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,ImBourreLaDoncSurLesComptes('44571'),PlanInfo.TVAAReverser,false,NivDetail);
      //EPZ 28/11/00
    end;
     *)
  end ;

  // BTY 10/05 En CRC 2002-10, dérogatoire supplémentaire si modif base
  // ------------------------------------------------------------------
  // fq 17773 - on ne fait rien en rapprochement bancaire - mbo 10.04.02006
  //if bModifBase then

  if (bModifBase) and (not bRappCpta) then
    begin
    ParamEcr.Methode:='NUL';
    if DeltaReprise > 0 then
      begin
      MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.Amort,DeltaReprise,True,bRappCpta,NivDetail,False);
      MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.Derog, DeltaReprise,False,bRappCpta,NivDetail,False);
      end
    else
      begin
      MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.Amort,(-1)*DeltaReprise,False,bRappCpta,NivDetail,False);
      MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.Derog, (-1)*DeltaReprise,True,bRappCpta,NivDetail,False);
      end;
    end;

  // BTY 11/05 Dérogatoire supplémentaire si révision du plan avec gestion d'un plan fiscal CRC 2002-10 pur
  // ------------------------------------------------------------------
  // fq 17773 - on ne fait rien en rapprochement bancaire - mbo 10.04.02006
  // if bRevPlanCRC then

  if (bRevPlanCRC) and (not bRappCpta) then
    begin
    ParamEcr.Methode:='NUL';
    if DeltaPlanCRC > 0 then
      begin
      MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.Amort,DeltaPlanCRC,True,bRappCpta,NivDetail,False);
      MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.Derog, DeltaPlanCRC,False,bRappCpta,NivDetail,False);
      end
    else
      begin
      MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.Amort,(-1)*DeltaPlanCRC,False,bRappCpta,NivDetail,False);
      MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CAss.Derog, (-1)*DeltaPlanCRC,True,bRappCpta,NivDetail,False);
      end;
    end;

  // PGR 12/2005 - Si Chgt méthode(Changement des conditions d'amortissement),
  //             rétablissement des amortissements antérieurs économiques et dérogatoires

  // fq 17773 - on ne fait rien en rapprochement bancaire - mbo 10.04.2006
  if (bChgtMethode) and (not bRappCpta) then
  begin
    ParamEcr.Methode:='NUL';

    if DeltaEco <> 0 then
    begin
      if DeltaEco > 0 then
      begin
        MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr, Q.FindField('I_COMPTEAMORT').AsString,DeltaEco,False,bRappCpta,NivDetail,True);
        MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,GetParamSocSecur('SO_IMMOANOUVEAU',''), DeltaEco,True,bRappCpta,NivDetail,True);
      end
      else
      begin
        MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,Q.FindField('I_COMPTEAMORT').AsString,(-1)*DeltaEco,True,bRappCpta,NivDetail,True);
        MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,GetParamSocSecur('SO_IMMOANOUVEAU',''), (-1)*DeltaEco,False,bRappCpta,NivDetail,True);
      end;
    end;

    if DeltaFisc <> 0 then
    begin
      if DeltaFisc > 0 then
      begin
        MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,Q.FindField('I_COMPTEDEROG').AsString,DeltaFisc,True,bRappCpta,NivDetail,True);
        MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,GetParamSocSecur('SO_IMMOANOUVEAU',''),DeltaFisc,False,bRappCpta,NivDetail,True);
      end
      else
      begin
        MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,Q.FindField('I_COMPTEDEROG').AsString,(-1)*DeltaFisc,False,bRappCpta,NivDetail,True);
        MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,GetParamSocSecur('SO_IMMOANOUVEAU',''), (-1)*DeltaFisc,True,bRappCpta,NivDetail,True);
      end;
    end;

    // Génération écritures inversées si sortie de l'immo
    if bSortie then
    begin
      if DeltaEco <> 0 then
      begin
        if DeltaEco > 0 then
        begin
          MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr, Q.FindField('I_COMPTEAMORT').AsString,DeltaEco,True,bRappCpta,NivDetail,True);
          MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,GetParamSocSecur('SO_IMMOANOUVEAU',''), DeltaEco,False,bRappCpta,NivDetail,True);
        end
        else
        begin
          MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,Q.FindField('I_COMPTEAMORT').AsString,(-1)*DeltaEco,False,bRappCpta,NivDetail,True);
          MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,GetParamSocSecur('SO_IMMOANOUVEAU',''), (-1)*DeltaEco,True,bRappCpta,NivDetail,True);
        end;
      end;

      if DeltaFisc <> 0 then
      begin;
        if DeltaFisc > 0 then
        begin
          MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,Q.FindField('I_COMPTEDEROG').AsString,DeltaFisc,False,bRappCpta,NivDetail,True);
          MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,GetParamSocSecur('SO_IMMOANOUVEAU',''),DeltaFisc,True,bRappCpta,NivDetail,True);
        end
        else
        begin
          MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,Q.FindField('I_COMPTEDEROG').AsString,(-1)*DeltaFisc,True,bRappCpta,NivDetail,True);
          MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,GetParamSocSecur('SO_IMMOANOUVEAU',''), (-1)*DeltaFisc,False,bRappCpta,NivDetail,True);
        end;
      end;
    end;
  end;

  // PGR - 12/2005    - Opération de Changement de méthode (Eclatement préalable) ACQ dans l'exercice
  // fq 17773 - on ne fait rien en rapprochement bancaire - mbo 10.04.2006
  if (bEclatPrealable) and (not bRappCpta) then
  begin
    ParamEcr.Methode:='NUL';
    MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr, Q.FindField('I_COMPTEIMMO').AsString,DeltaHt,True,bRappCpta,NivDetail,True);
    MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,GetParamSocSecur('SO_IMMOANOUVEAU',''), DeltaHt,False,bRappCpta,NivDetail,True);

    // Génération écritures inversées si sortie de l'immo
    if bSortie then
    begin
      MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr, Q.FindField('I_COMPTEIMMO').AsString,DeltaHt,False,bRappCpta,NivDetail,True);
      MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,GetParamSocSecur('SO_IMMOANOUVEAU',''), DeltaHt,True,bRappCpta,NivDetail,True);
    end;
  end;

  //mbo - 02.11.06 - Saisie d'une subvention d'investissement
  if Plan.SBV then
  begin
    PlanInfo := TPlanInfo.Create(ImmoEnCours);
    try
      PlanInfo.ChargeImmo(ImmoEnCours);
      PlanInfo.Calcul(ParamEcr.DateCalcul, true, false, ''); //fq 21754

      if bRappCpta then  // rapprochement bancaire
      begin
        fprorata:=false;
        ParamEcr.Methode:='LIN';

        // MVG 05/12/2006 Réduction compte 131 et 77
        if bReducSBV then
        begin
          //MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CSbv.Dotation,DeltaReduc,true,bRappCpta,NivDetail,False);
          // FQ 19281 MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CSbv.Reprise,DeltaReduc,false,bRappCpta,NivDetail,False);
          MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CSbv.Quotepart,DeltaReduc,false,bRappCpta,NivDetail,False);
        end;

        if (not bsortie) and (not bCessionTotale) then
        begin
          // MVG 05/12/2006 le compte 131 est de sens crédit
//          MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CSbv.Dotation,PlanInfo.BaseSBV,true,bRappCpta,0,False);
          MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CSbv.Dotation,PlanInfo.BaseSBV,false,bRappCpta,0,False);
          MajLigneEcriture(L,Indice,ImmoEnCours,ParamEcr,CSbv.Reprise,PlanInfo.CumulSbv,true,bRappCpta,0, False);
          // MVG Mauvais montant que le compte 777
//        MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CSbv.QuotePart,PlanInfo.VNCSbv,false,bRappCpta,0,False);
          MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CSbv.QuotePart,PlanInfo.DotationSbv,false,bRappCpta,0,False);
       end else
        begin
//           MajLigneEcriture(L,Indice,ImmoEnCours,ParamEcr,CSbv.QuotePart,PlanInfo.VNCSbv,false,bRappCpta,0, False);
//           MVG Mauvais montant que le compte 777
             MajLigneEcriture(L,Indice,ImmoEnCours,ParamEcr,CSbv.QuotePart,PlanInfo.DotationSbv,false,bRappCpta,0, False);
        end;
      end else
      begin // génération des écritures
        fprorata := true;
        ParamEcr.Methode := 'NUL';
        // On génére le jeu d'écriture lié à la réduction de subvention si besoin
        if bReducSBV then
        begin
          MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CSbv.Dotation,DeltaReduc,true,bRappCpta,NivDetail,False);
          // FQ 19282 MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CSbv.Reprise,DeltaReduc,false,bRappCpta,NivDetail,False);
          MajLigneEcriture (L,Indice,ImmoEnCours,ParamEcr,CSbv.Quotepart,DeltaReduc,false,bRappCpta,NivDetail,False);
        end;
        // on génére les jeux d'écritures à la subvention proprement dite (dotation ou vnc si sortie
        if bSortie and bCessionTotale then
        begin
          MajLigneEcriture(L,Indice,ImmoEnCours,ParamEcr,CSbv.Reprise,PlanInfo.DotationSbv,true,bRappCpta,NivDetail, False);
          MajLigneEcriture(L,Indice,ImmoEnCours,ParamEcr,CSbv.QuotePart,PlanInfo.DotationSbv,false,bRappCpta,NivDetail, False);

          MajLigneEcriture(L,Indice,ImmoEnCours,ParamEcr,CSbv.Dotation,PlanInfo.BaseSBV,true,bRappCpta,NivDetail, False);
          MajLigneEcriture(L,Indice,ImmoEnCours,ParamEcr,CSbv.Reprise,PlanInfo.BaseSBV,false,bRappCpta,NivDetail, False);
        end else
        begin
          MajLigneEcriture(L,Indice,ImmoEnCours,ParamEcr,CSbv.Reprise,PlanInfo.DotationSbv,true,bRappCpta,NivDetail, False);
          MajLigneEcriture(L,Indice,ImmoEnCours,ParamEcr,CSbv.QuotePart,PlanInfo.DotationSbv,false,bRappCpta,NivDetail, False);
        end;

      end;
    finally
      PlanInfo.Free;
    end;
  fprorata:=true;
  end;
  Plan.Free ;
end;


////////////////////////////////////////////////////////////////////////////////
{

   Fonctions & procedures de génération des écritures de dotation & échéances

}
procedure CalculEcrituresDotation ( var L : TList;Code : string;ParamEcr : TParamEcr;NivDetail : integer);
var Q : TQuery; sSelect , sWhere : string;
begin
//  CurPiece := '';
  ListeCompte := TList.Create;
  try
    if Code = '' then // Toutes les écritures de dotation
    begin
      sSelect := 'SELECT * FROM IMMO';
      sWhere  := ' WHERE I_NATUREIMMO="PRO" AND I_ETAT<>"FER" AND I_QUALIFIMMO="R"' ;
      (*if NivDetail = 4 then sWhere := ' WHERE I_NATUREIMMO="PRO" AND I_ETAT<>"FER" AND I_QUALIFIMMO="R"'
                         else sWhere := ' WHERE I_NATUREIMMO="PRO" AND I_ETAT<>"FER" AND I_QUALIFIMMO="R"';*)
    end
    else  // Ecritures de dotation pour une Immo
    begin
      sSelect := 'SELECT * FROM IMMO';
      if NivDetail = 4 then sWhere := ' WHERE I_IMMO="'+Code+'" AND I_ETAT<>"FER" AND I_QUALIFIMMO="R"'
                       else sWhere := ' WHERE I_IMMO="'+Code+'" AND I_NATUREIMMO="PRO" AND I_ETAT<>"FER" AND I_QUALIFIMMO="R"';
    end;
    Q := OpenSQL (sSelect+sWhere,TRUE);
    try
      {$IFDEF EAGLCLIENT}
      InitMove(Q.Detail.Count,'');
      {$ELSE}
      InitMove(QCount(Q),'');
      {$ENDIF}
      Dotation_V2 := False;
      while not Q.Eof do
      begin
        CalculIntegrationDotation (Q,L,ParamEcr,NivDetail);
        MoveCur(False);
        Q.Next;
      end;
    finally
      Ferme (Q);
      FiniMove;
    end ;
  finally
    VideListeCompte(ListeCompte) ; // DetruitListeComptes (ListeCompte);
    ListeCompte.Free ;
  end ;
end;

////////////////////////////////////////////////////////////////////////////////
procedure CalculEcrituresEcheances ( var L : TList;Code : string;ParamEch,ParamPai : TParamEcr;NivDetail : integer; dtEchMin, dtEchMax : TDateTime;AvecDejaIntegre: boolean=false);
var Q : TQuery; LeSql: string ;
begin
//  CurPiece := '';
  ListeCompte := TList.Create;
  LeSql:='SELECT IMMO.*,G_TVA FROM IMMO LEFT JOIN GENERAUX ON (G_GENERAL=I_COMPTELIE) '
        +' WHERE (I_NATUREIMMO="CB" OR I_NATUREIMMO="LOC") AND I_ETAT<>"FER" AND I_QUALIFIMMO="R"';
  if Code<>'' then LeSql:=LeSql+' AND I_IMMO="'+Code+'" ';
  Q:=OpenSQL (LeSql,TRUE);
  if not Q.Eof then
  begin
    Q.First;
    while not Q.Eof do
    begin
      CalculIntegrationEcheance (Q,L,ParamEch,ParamPai,NivDetail,dtEchMin,dtEchMax,AvecDejaIntegre);
      MoveCur(False);
      Q.Next;
    end;
  end;
  VideListeCompte(ListeCompte) ; // DetruitListeComptes (ListeCompte);
  ListeCompte.Free ;
  Ferme(Q) ;
end;


////////////////////////////////////////////////////////////////////////////////
procedure VideListeEcritures(L : TList);
var i:integer;
    ARecord:TLigneEcriture;
    {$IFDEF SERIE1}
    {$ELSE}
    j,k:integer;
    ARecordAna : TAna;
    {$ENDIF}
begin
  if L=Nil then Exit ; if L.Count<=0 then Exit ;
  for i:=0 to (L.Count - 1) do
  begin
    ARecord:=L.Items[i];
    {$IFDEF SERIE1}
    {$ELSE}
    for j:=0 to ImMaxAxe-1 do
    begin
      if (*Pos(InttoStr(i),'12345')>0 and *)ARecord.Axes[j]<>nil then   //??? YCP
      begin
        for k := 0 to (ARecord.Axes[j].Count - 1) do
        begin
          ARecordAna := ARecord.Axes[j].Items[k];
          ARecordAna.Free;
        end;
        ARecord.Axes[j].Free;
        ARecord.Axes[j]:=nil;
      end;
    end;
    {$ENDIF}
    ARecord.Free ;
  end;
  L.Clear ;
end;

////////////////////////////////////////////////////////////////////////////////
// Fontion de recalcul du montant d'une dotation en fonction de la date de calcul des dotations
function RecalculMontant (Valeur : double; DateCalcul,DateAmort,DateFinAmort : TDateTime;bDotation : boolean;Methode : string; AvecSurAMort : boolean) : double;
var Prorata : double;
    dtDebutCalc : TDateTime;
begin
  if Dotation_V2 then
  begin
    Result := Valeur;
    exit;
  end;
  if (DateCalcul = VHImmo^.Encours.Fin) or (not bDotation) or (Methode = 'NUL') then Result := Valeur else
  begin
    if (DateFinAmort < VHImmo^.Encours.Deb) or (DateFinAmort > VHImmo^.Encours.Fin) then
     DateFinAmort := iDate1900;
    if (DateAmort > VHImmo^.Encours.Deb) and (DateAmort < VHImmo^.Encours.Fin) then dtDebutCalc := DateAmort
    else dtDebutCalc := VHImmo^.Encours.Deb;
    // ajout mbo - FQ 20841 - 25.06.07 pour gestion du plan var qui n'était pas prévue
    if (Methode = 'LIN') or (Methode = 'VAR') then
      result := GetDotationLinAvant(dtDebutCalc,VHImmo^.Encours.Fin,DateAmort,DateCalcul,DateFinAmort,Valeur,False)
    else if Methode = 'DEG' then
    begin
      if (DateFinAmort < VHImmo^.Encours.Deb) or (DateFinAmort > VHImmo^.Encours.Fin) then
         DateFinAmort := VHImmo^.Encours.fin;
      result := GetDotationDegAvant(dtDebutCalc,DateFinAmort,DateAmort,DateCalcul , Valeur,Prorata, AvecSurAmort);
    end
    else result := 0.0;
  end;
end;

{ TLigneEcriture }
procedure TLigneEcriture.Copy(LigneEcr: TLigneEcriture);
{$IFDEF SERIE1}
{$ELSE}
var i,j : integer; ARecord, ARecordNew : TAna;
{$ENDIF}
begin
  CodeImmo := LigneEcr.CodeImmo;
  Piece := LigneEcr.Piece;
  Compte := LigneEcr.Compte;
  CompteRef := LigneEcr.CompteRef;
  Auxi := LigneEcr.Auxi;
  Libelle := LigneEcr.Libelle;
  Debit := LigneEcr.Debit;
  Credit := LigneEcr.Credit;
  Date := LigneEcr.Date;
  CodeJournal := LigneEcr.CodeJournal;
  Couleur := LigneEcr.Couleur;
  Etablissement := LigneEcr.Etablissement;
  {$IFDEF SERIE1}
  {$ELSE}
  for i:=1 to ImMaxAxe do
  begin
    if LigneEcr.Axes[i-1]<>nil then
    begin
      Axes[i-1] := TList.Create;
      for j:=0 to LigneEcr.Axes[i-1].Count - 1 do
      begin
        ARecord := LigneEcr.Axes[i-1].Items[j];
        ARecordNew := TAna.Create;
        ARecordNew.Copy (ARecord);
        Axes[i-1].Add (ARecordNew);
      end;
    end;
  end;
  {$ENDIF}
end;

{ TAna }

{$IFDEF SERIE1}
{$ELSE}
procedure TAna.Copy(ARecord: TAna);
begin
    NumeroVentil := ARecord.NumeroVentil;
    Section := ARecord.Section;
    Libelle := ARecord.Libelle;
    TauxMontant := ARecord.TauxMontant;
    TauxQte1 := ARecord.TauxQte1;
    TauxQte2 := ARecord.TauxQte2;
    Montant := ARecord.Montant;
end;
{$ENDIF}

{ TAmListeEcriture }
{$IFDEF PLUS_TARD}

procedure TAmListeEcriture.AjouteLigne(DateComptable : TDateTime; Compte: string; Debit, Credit: double);
var
  AmEcr : TLigneEcriture;
begin
  AmEcr := nil;
  if ( Compte = '' ) or ((Debit = 0) and (Credit = 0)) then exit;
  case fNiveauDetail of
    ndSynth : AmEcr := TrouveEcriture ( DateComptable, Compte );
    ndCompte : AmEcr := TrouveEcriture ( DateComptable, Compte );
  end;
  if (AmEcr <> nil) then
  begin
//    AmEcr.Debit
    Montant:=RecalculMontant(Montant,ParamEcr.DateCalcul,ParamEcr.DateAmort,ParamEcr.DateFinAmort,ParamEcr.bDotation,ParamEcr.Methode, AvecSurAmort) ;
    if Debit then ARecord.Debit :=ARecord.Debit +Montant else ARecord.Credit:=ARecord.Credit+Montant ;
    {$IFDEF SERIE1}
    {$ELSE}
    if not bRappCpta then ChargeVentilImo (IsCompteVentilable(ARecord.Compte),ARecord.Axes,'IM',CodeImmo,Montant,ARecord.Debit+ARecord.Credit) ;
    {$ENDIF}
  end
  else
  begin
    ARecord := TLigneEcriture.Create;
    if ParamEcr.bDotation then
    case NivDetail of
      0: ARecord.Piece := '';
      1: ARecord.Piece := ParamEcr.CompteRef;
      2: ARecord.Piece := CodeImmo;
    end;
    if not ParamEcr.bDotation then ARecord.Piece := IntToStr(indice);
//    CurPiece := ARecord.Piece;
    ARecord.CodeImmo := CodeImmo;
    ARecord.Credit := 0.0; ARecord.Debit := 0.0;
    ARecord.Date := ParamEcr.Date;
    ARecord.CodeJournal := ParamEcr.Journal;
    if NivDetail = 1 then ARecord.CompteRef := ParamEcr.CompteRef
    else ARecord.CompteRef := '';
    ARecord.Etablissement := ParamEcr.Etablissement;
    ARecord.Libelle := RecupereLibelle(Compte,bGeneral);
    if bGeneral then ARecord.Compte := Compte else ARecord.Auxi := Compte;
    if ParamEcr.Libelle <> '' then ARecord.Libelle := ParamEcr.Libelle;
    if Debit then ARecord.Debit :=RecalculMontant(Montant,ParamEcr.DateCalcul,ParamEcr.DateAmort,ParamEcr.DateFinAmort,ParamEcr.bDotation,ParamEcr.Methode, AvecSurAmort)
             else ARecord.Credit:=RecalculMontant(Montant,ParamEcr.DateCalcul,ParamEcr.DateAmort,ParamEcr.DateFinAmort,ParamEcr.bDotation,ParamEcr.Methode, AvecSurAmort);
    {$IFDEF SERIE1}
    {$ELSE}
    if not bRappCpta then ChargeVentilImo (IsCompteVentilable(Compte),ARecord.Axes,'IM',CodeImmo,Montant,ARecord.Debit+ARecord.Credit);
    {$ENDIF}
    L.Add(ARecord);
  end;
end;
  {$ENDIF}

end.
