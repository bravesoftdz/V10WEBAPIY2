unit FactCptaMode;

interface
Uses UTOB,
{$IFDEF EAGLCLIENT}
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     HCtrls,SysUtils,HEnt1,UtilPGI,Ent1,EntGC,FactCpta;
     
Function RenseigneTobEchesCptaFFO ( TOBPiece,TOBEchesCpta : TOB ) : integer;
Function EnregistrePiedEcheFFO ( TOBPiece,TOBTiers,TOBArticles,TOBEches : TOB; LeAcc : T_GCAcompte ) : Boolean ;
Function TestReglementArticleFi ( TOBEche : TOB ) : boolean ;

implementation

// Regroupement des lignes de la Tob Echeance, qui ont les mêmes modes de paiement,
// date d'échéance et devise. Cas d'un paiement en Espèces avec rendu-monnaie.
Procedure CumulLignesEcheance (TobEche1 : TOB ) ;
Var TobEche2, T1, T2 : TOB ;
BEGIN
if TobEche1.Detail.Count > 1 then
   begin
   TobEche2:=TOB.Create('Echeance regroupee',Nil,-1) ;
   T1:=TobEche1.Detail[0];
   T1.ChangeParent (TobEche2, -1);
   While TobEche1.Detail.Count > 0 do
      begin
      T1:=TobEche1.Detail[0];
      T2 := TobEche2.FindFirst (['GPE_MODEPAIE','GPE_DATEECHE','GPE_DEVISE'], [T1.GetValue('GPE_MODEPAIE'),T1.GetValue('GPE_DATEECHE'),T1.GetValue('GPE_DEVISE')], False);
      if T2 <> Nil then
         begin
         T2.PutValue ('GPE_MONTANTECHE',T2.GetValue('GPE_MONTANTECHE')+T1.GetValue('GPE_MONTANTECHE'));
         T2.PutValue ('GPE_MONTANTDEV',T2.GetValue('GPE_MONTANTDEV')+T1.GetValue('GPE_MONTANTDEV'));
         T2.PutValue ('GPE_MONTANTENCAIS',T2.GetValue('GPE_MONTANTENCAIS')+T1.GetValue('GPE_MONTANTENCAIS'));
         T1.Free;
         end
         else T1.ChangeParent (TobEche2, -1);
      end;
   // Réaffectation de l'ensemble des lignes sur la Tob Echéance initiale.
   While TobEche2.Detail.Count > 0 do
      begin
      T2:=TobEche2.Detail[0];
      T2.ChangeParent (TobEche1, -1);
      end;
   TobEche2.free;
   end;
END;

// Génération de lignes d'Echéance distinctes par article financier.
Procedure LignesEcheanceArtFi (TobArtFinancier, TobEche1 : TOB ) ;
Var TobEche2, T1, T2 : TOB ;
    TobArtF : TOB ;
    MontantEche,MontantArtF,ResteARegler : Double ;
    Montant, MontantDev : Double ;
    Devise : String ;
 //   DateConv : TDateTime;
    i, ie : Integer ;
BEGIN
ResteARegler:=0;
TobEche1.Detail[0].AddChampSup('REGLEFI', True); // Règlement d'un article financier ? (X ou -)
TobEche1.Detail[0].AddChampSup('ARTICLE', True); // Pour stocker le code article financier
TobEche2:=TOB.Create('Echeance ArtFi',Nil,-1) ;
for i:=0 to TobArtFinancier.Detail.Count-1 do
  begin
  TobArtF := TobArtFinancier.Detail[i] ;
  MontantArtF := TobArtF.GetValue('MONTANTTC');
  While TobEche1.Detail.Count > 0 do
    begin
    T1:=TobEche1.Detail[0];
    T2:=TOB.Create('PIEDECHE',TobEche2,-1) ;
    T2.Dupliquer(T1,False,True) ;
    MontantEche:=T1.GetValue('GPE_MONTANTDEV');
    if MontantArtF = MontantEche then
       begin
       // Affectation du total du règlement à l'article financier
       ResteARegler := 0;
       T1.Free;
       end;
    if MontantArtF > MontantEche then
       begin
       // Affectation du total du règlement à l'article financier
       // Le solde du montant de l'article financier sera affecté sur la ligne
       //    d'échéance suivante.
       ResteARegler := MontantArtF-MontantEche;
       T1.Free;
       end;
    if MontantArtF < MontantEche then
       begin
       ResteARegler := 0;
       // Affectation d'une partie du règlement à l'article financier
       Montant    := 0 ;
       MontantDev := MontantArtF ;
       Devise     := T2.GetValue('GPE_DEVISE');
 //      Dateconv   := NowH ;
       T2.PutValue ('GPE_MONTANTDEV',  MontantDev) ;
       T2.PutValue ('GPE_MONTANTECHE', Montant) ;

       // Le reste du règlement apparaîtra sur une nouvelle ligne d'échéance.
       Montant    := 0 ;
       MontantDev := MontantEche-MontantArtF ;
       T1.PutValue('GPE_MONTANTDEV', MontantDev) ;
       T1.PutValue('GPE_MONTANTECHE',Montant) ;
       end;
    // Mémorisation du code Article financier dans la Tob Echeance
    T2.PutValue('REGLEFI','X');
    T2.PutValue('ARTICLE',TobArtF.GetValue('ARTICLE'));
    if ResteARegler = 0 then break;
    MontantArtF := ResteARegler;
    end;
  end;

// Recopie des lignes d'échéances suivantes, non affectées à des produits financiers
While TobEche1.Detail.Count > 0 do
  begin
  T1:=TobEche1.Detail[0];
  T1.ChangeParent (TobEche2, -1);
  end;

// Réaffectation de l'ensemble des lignes sur la Tob Echéance de départ
While TobEche2.Detail.Count > 0 do
  begin
  T2:=TobEche2.Detail[0];
  T2.ChangeParent (TobEche1, -1);
  end;

ie:=0;  
for i:=0 to TobEche1.Detail.Count-1 do
    BEGIN
    if i=0 then ie:=TobEche1.Detail[i].GetNumChamp('GPE_NUMECHE') ;
    TobEche1.Detail[i].PutValeur(ie,i+1) ;
    END ;

TobEche2.Free;
END;

// ---------------------------------------------------------------------------
// Génération d'une Tob des échéances pour la Compta, distinguant les montants
// de règlement par Article financier.
// ---------------------------------------------------------------------------
Function RenseigneTobEchesCptaFFO (TOBPiece,TOBEchesCpta : TOB ) : integer;
Var TobL, TobArtF, TobArtFinancier : TOB ;
    NbArtF, NbTotArtF : Integer ;
    MontantArtF : Double ;
BEGIN
Result:=0;
if TOBEchesCpta.Detail.Count=0 then Exit ;

// Regroupement des lignes de règlement FFO ayant les mêmes mode de paiement,
// date échéance et devise. Cas d'un paiement en espèces avec rendu monnaie.
CumulLignesEcheance (TOBEchesCpta);

// Recherche si des produits financiers ont été saisis sur le ticket FFO
NbArtF:=0; NbTotArtF:=0;
TobL:=TOBPiece.FindFirst (['GL_TYPEARTICLE'], ['FI'], False);
if TobL<>Nil then
   begin
   TobArtFinancier:= TOB.CREATE ('Art-Financier', nil, -1);
   While TobL <> Nil do
      begin
      MontantArtF := TobL.GetValue ('GL_MONTANTTTCDEV');
      if MontantArtF <> 0 then
         begin
         TobArtF:= TOB.CREATE ('Art-F', TobArtFinancier, -1);
         TobArtF.AddChampSup ('ARTICLE', False);
         TobArtF.AddChampSup ('MONTANTTC', False);
         TobArtF.PutValue    ('ARTICLE', TobL.GetValue ('GL_ARTICLE'));
         TobArtF.PutValue    ('MONTANTTC', TobL.GetValue ('GL_MONTANTTTCDEV'));
         inc(NbArtF );
         end;
      inc(NbTotArtF );
      TobL:=TOBPiece.FindNext (['GL_TYPEARTICLE'], ['FI'], False);
      end;
   // Des articles financiers sont enregistrés sur le ticket de vente, on génère
   // une ligne de règlement distincte par article financier.
   if NbArtF<>0 then LignesEcheanceArtFi (TobArtFinancier, TOBEchesCpta ) ;
   TobArtFinancier.Free;
   end;
Result:=NbTotArtF;
END ;


// ----------------------------------------------------------------------------
// Génération d'écritures d'Acomptes/Règlements pour la nature de document FFO,
// à partir de la table PIEDECHE.
// ----------------------------------------------------------------------------
Function EnregistrePiedEcheFFO ( TOBPiece,TOBTiers,TOBArticles,TOBEches : TOB; LeAcc : T_GCAcompte ) : Boolean ;
Var Modepaie,CodJrl : String ;
    i : integer ;
    Q : TQuery ;
    TOBE,TobAc,TobArt : TOB ;
    SQL  : string ;
  //  XD,XP,XE : Double ;
BEGIN
Result:=True;
for i:=0 to TOBEches.Detail.Count-1 do
    BEGIN
    TOBE := TOBEches.Detail[i] ;
    Modepaie:=TOBE.GetValue('GPE_MODEPAIE');
    FillChar(LeAcc,Sizeof(LeAcc),#0) ;
    SQL:='SELECT MP_JALREGLE, MP_CPTEREGLE, MP_LIBELLE From MODEPAIE Where MP_MODEPAIE="'+Modepaie+'"';
    Q := OpenSQL(SQL, true,-1, '', True);
    if not Q.EOF then
       begin
       LeAcc.ModePaie := ModePaie;
       LeAcc.JalRegle := Q.FindField('MP_JALREGLE').AsString;
       LeAcc.CpteRegle := Q.FindField('MP_CPTEREGLE').AsString;
       LeAcc.Libelle := Q.FindField('MP_LIBELLE').AsString;
       end else
       begin
       LeAcc.JalRegle := '';
       LeAcc.CpteRegle := '';
       end ;
    Ferme(Q);
    if (LeAcc.JalRegle <> '') and (LeAcc.CpteRegle <> '') then
       begin
       if TOBE.GetValue('GPE_DEVISE')<>V_PGI.DevisePivot
        then LeAcc.Montant:=TOBE.GetValue('GPE_MONTANTDEV')
        else LeAcc.Montant:=TOBE.GetValue('GPE_MONTANTECHE');
       LeAcc.DateEche:=DateToStr(TOBE.GetValue('GPE_DATEECHE'));
       LeAcc.IsReglement:=true ;
       LeAcc.IsContrepartie:=false;
       LeAcc.CpteContre:='';
       LeAcc.LibelleContre:='';
       if (TestReglementArticleFi(TOBE)=true) then
          begin
          // Règlement d'un article de type financier
          // Recherche du compte rattaché à cet article, qui serait un compte différent du compte de Tiers
          TobArt:=TOBArticles.FindFirst(['GA_ARTICLE'],[TOBE.GetValue('ARTICLE')],False) ;
          if TobArt<>Nil then
             begin
             LeAcc.CpteContre:=Trim(Copy(TobArt.GetValue('GA_DESIGNATION1'),1,17));
             LeAcc.LibelleContre:=Trim(Copy(TobArt.GetValue('GA_LIBELLE'),1,35));
             LeAcc.Libelle:=LeAcc.LibelleContre;
             if LeAcc.CpteContre<>'' then LeAcc.IsContrepartie:=true;
             // Reprise du code Journal personnalisé pour l'article financier, si celui-ci est renseigné
             CodJrl:=Trim(Copy(TobArt.GetValue('GA_DESIGNATION2'),1,3));
             if CodJrl<>'' then LeAcc.JalRegle:=CodJrl;
             end;
          end ;
       if (LeAcc.CpteRegle <> LeAcc.CpteContre) then
          begin
          TobAc:=EnregistreAcompte(TOBPiece,TOBTiers,LeAcc) ;
          if (TobAc <> Nil) then TobAc.free ;
          end;
       end ;
    END ;
END ;

// Test si l'échéance (Règlement FFO) est liée à un article de type financier
Function TestReglementArticleFi ( TOBEche : TOB ) : boolean ;
BEGIN
Result:=False ;
if (TOBEche.FieldExists('REGLEFI')=true) and (TOBEche.GetValue('REGLEFI')='X') then Result:=True;
END ;

end.
