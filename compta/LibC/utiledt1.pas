unit UtilEdt1 ;

interface

Uses UtilEdt,CritEdt,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     ENT1,HENT1,HCTRLS,HQRY,STDCTRLS,SYSUTILS,MASK,
     CLASSES,Hcompte,Forms, HQuickrp,graphics   ,UentCommun;

Const OkV2 : Boolean = FALSE ;

Type TTabDate4 = Array[1..4] of TDateTime ;
Type TTabDate8 = Array[1..8] of TDateTime ;
Type TTabDate12 = Array[1..12] of TDateTime ;
Type TTabDate24 = Array[1..24] of TDateTime ;
type  TModPaie = Class
        Montant : Double ;
        END ;

type TFTotAnaGene = Class
    Sect,Gene : String17 ;
    TotAG : TabMont77 ;
    END ;

Procedure genereSQLTiers(NE : TNatureEtat ; NET : TNatureEtatTiers ; Var CritEdt : TCritEdt ; Q : TQuery ; OkTP : Boolean ; EtatRevision : TCheckBoxState) ;
Procedure genereSQLTiersMultiDevise(NE : TNatureEtat ; NET : TNatureEtatTiers ; Var CritEdt : TCritEdt ; Q : TQuery) ;
Procedure genereSQLSUBTiers(NE : TNatureEtat ; NET : TNatureEtatTiers ; Var CritEdt : TCritEdt ; QEcr : TQuery ; OkTP : Boolean) ;
procedure CalculDateTiers(NE : TNatureEtat ; NET : TNatureEtatTiers ; Var CritEdt : TCritEdt ;
                          Var TabDate : TTabDate4 ;
                          FP1,FP2,FP3,FP4,FP5,FP6,FP7,FP8 : TMaskEdit) ;
Function ValidPeriodeTiers(FP1,FP2,FP3,FP4,FP5,FP6,FP7,FP8 : TMaskEdit) : Boolean ;
procedure ChargeRecapMPTiers (Var L : TStringList; MultiCombo : THMultiValcomboBox) ;
function MDPRetenuTiers(Var L : TStringList ; St : string) : Boolean ;
procedure PeriodiciteChangeTiers(NE : TNatureEtat ; NET : TNatureEtatTiers ; Iperiodicite : Integer ;
                                 FP8,FP1 : TMaskEdit ; Retard : Boolean ; Var TabD : TTabDate8) ;

{ Rony 18/04/97 -- Change ZoomTable des colls pour
                   GlVen, BLVen, GlAge, BlAge
                   Prevoir aussi pour les autres editions, donc utiliser NE, NET }
Procedure ChangeColl(Nature : String ; C1, C2 : THCpteEdit) ;

{ Retaillage des bandes en Edition Budgetaire }
Procedure RetailleBandesBudget(FF : TForm ; CritEdt : TCritEdt ; LibRev : String ; BalOuEca : Boolean ; BD, BDLibre, BDFin : TQRBAND) ;

Procedure RetailleBandesBudgetes(FF : TForm ; CritEdt : TCritEdt ; LibRev : String ; BD, BDLibre, BDRecap, BDFPrim, BDFin : TQRBAND) ;
Procedure RetailleBandesBudEcarts(FF : TForm ; CritEdt : TCritEdt ; LibRev : String ; BD, BDLibre, BDRecap, BDFin : TQRBAND) ;

implementation

Uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  CpteUtil ;

Procedure genereSQLTiers(NE : TNatureEtat ; NET : TNatureEtatTiers ; Var CritEdt : TCritEdt ; Q : TQuery ; OkTP : Boolean ; EtatRevision : TCheckBoxState) ;
Var St,St1,StSup : String ;
BEGIN
Case CritEdt.GlV.TypePar of
  0 : BEGIN     { Tiers demandé }
      Q.Close ; Q.SQL.Clear ;
      Q.SQL.Add(' Select T_AUXILIAIRE , T_LIBELLE, T_CREDITACCORDE, T_ISPAYEUR as ISPAYEUR, T_PAYEUR AS PAYEUR ') ;
      Case CritEdt.Rupture of
        rLibres  : Q.SQL.Add(', T_TABLE0, T_TABLE1, T_TABLE2, T_TABLE3, T_TABLE4, T_TABLE5, T_TABLE6, T_TABLE7, T_TABLE8, T_TABLE9 ') ;
        rCorresp : Q.Sql.Add(', T_CORRESP1, T_CORRESP2 ') ;
        End ;

      Q.SQL.Add(' From TIERS T Where ') ;
      { Construction de la clause Where de la SQL }
      if CritEdt.DeviseSelect<>'' then
         BEGIN
         Q.SQL.Add(' (Exists (Select E_AUXILIAIRE,E_DATECOMPTABLE From ECRITURE Where E_AUXILIAIRE=T.T_AUXILIAIRE ') ;
         Q.SQL.Add(' And E_DEVISE="'+CritEdt.DeviseSelect+'")) ') ;
         END else
         BEGIN
         Q.SQL.Add(' T_AUXILIAIRE <> "'+w_w+'" ') ;
         END ;
      If OkTP Then
         BEGIN
         Q.SQL.Add(' AND T_PAYEUR<>"" AND T_ISPAYEUR<>"X" ') ;
         if CritEdt.GlV.TP1<>'' then Q.SQL.Add(' AND T_PAYEUR>="'+CritEdt.GlV.TP1+'" ') ;
         if CritEdt.GlV.TP2<>'' then Q.SQL.Add(' AND T_PAYEUR<="'+CritEdt.GlV.TP2+'" ') ;
         END ;
      if CritEdt.Joker then Q.SQL.Add(' AND T_AUXILIAIRE like "'+TraduitJoker(CritEdt.Cpt1)+'" ') Else
         BEGIN
         if CritEdt.Cpt1<>'' then Q.SQL.Add(' AND T_AUXILIAIRE>="'+CritEdt.Cpt1+'" ') ;
         if CritEdt.Cpt2<>'' then Q.SQL.Add(' AND T_AUXILIAIRE<="'+CritEdt.Cpt2+'" ') ;
         END ;
      if CritEdt.NatureCpt<>'' then Q.SQL.Add(' AND T_NATUREAUXI="'+CritEdt.NatureCpt+'" ')
                               Else Q.SQL.Add(' AND T_NATUREAUXI<>"NCP" ') ;
      Q.SQL.Add(' AND T_LETTRABLE ="X" ') ;
//      Q.SQL.Add(' AND (T_TOTALDEBIT<>0 or T_TOTALCREDIT<>0) ') ;

      St1:=WhereSuppEdtTiers(CritEdt.QualifPiece,V_PGI.Controleur,EtatRevision) ;
      if VH^.ExoV8.Code<>'' then
         BEGIN
         Q.SQL.Add(' AND (Exists (Select E_AUXILIAIRE,E_ETATLETTRAGE From ECRITURE Where E_AUXILIAIRE=T.T_AUXILIAIRE ') ;
         If St1<>'' Then Q.SQL.Add(St1) ;
         //RR 18/07/2005 Attention, pb de récupération, modification de la requete afin de récupérer les Tiers facturé TOTALEMENT LETTREE
         if OkTP then
          Q.SQL.Add(' And E_ETATLETTRAGE="TL" And (E_DATECOMPTABLE>="'+UsDateTime(VH^.ExoV8.Deb)+'") AND (E_QUALIFPIECE<>"C" AND E_ECRANOUVEAU<>"OAN"))) ')
         else
          Q.SQL.Add(' And (E_ETATLETTRAGE="AL" OR E_ETATLETTRAGE="PL") And (E_DATECOMPTABLE>="'+UsDateTime(VH^.ExoV8.Deb)+'") AND (E_QUALIFPIECE<>"C" AND E_ECRANOUVEAU<>"OAN"))) ') ;
         END Else
         BEGIN
         Q.SQL.Add(' AND (Exists (Select E_AUXILIAIRE,E_ETATLETTRAGE From ECRITURE Where E_AUXILIAIRE=T.T_AUXILIAIRE ') ;
         If St1<>'' Then Q.SQL.Add(St1) ;
          //RR 18/07/2005 Attention, pb de récupération, modification de la requete afin de récupérer les Tiers facturé TOTALEMENT LETTREE
         if OkTP then
          Q.SQL.Add(' And E_ETATLETTRAGE="TL" And (E_DATECOMPTABLE>="'+UsDateTime(VH^.ExoV8.Deb)+'") AND (E_QUALIFPIECE<>"C" AND E_ECRANOUVEAU<>"OAN"))) ')
         else
          Q.SQL.Add(' And (E_ETATLETTRAGE="AL" OR E_ETATLETTRAGE="PL") And (E_QUALIFPIECE<>"C" AND E_ECRANOUVEAU<>"OAN"))) ') ;
         END ;
      If Not OkTP Then If CritEdt.GLV.SaufCptSolde Then Q.SQL.Add(' AND (T_TOTALDEBIT<>T_TOTALCREDIT) ') ;
// BPY le 15/01/2004 => fiche 12403 : ajout de la confidentialité ....
// => old one ...
//      if (OkV2 and (V_PGI.Confidentiel<>'1')) then Q.SQL.Add(' AND T_CONFIDENTIEL<>"1" ') ;
// => pour les future confidantialité !
//      if (V_PGI.Confidentiel = '1') then Q.SQL.Add(' AND T_CONFIDENTIEL<=' + IntToStr(V_PGI.NiveauAccesConf))
//      else Q.SQL.Add(' AND T_CONFIDENTIEL=0');
// => bah ...
      if (V_PGI.Confidentiel <> '1') then Q.SQL.Add(' AND T_CONFIDENTIEL <> "1" ');
// fin BPY
      { Construction de la clause Order By de la SQL }
      StSup:='' ; If OkTP Then StSup:='T_PAYEUR, ' ;
      Case CritEdt.Rupture of
        rLibres  : BEGIN
                   St:=WhereLibre(CritEdt.LibreCodes1,CritEdt.LibreCodes2,fbAux,CritEdt.GlV.OnlyCptAssocie) ;
                   If St<>'' Then Q.Sql.Add(st) ;
                   Q.SQL.Add('Order By '+StSup+OrderLibre(CritEdt.LibreTrie)+'T_AUXILIAIRE ') ;
                   END ;
        rCorresp : Case CritEDT.GLV.PlansCorresp Of
                     1 : Q.Sql.Add('Order By '+StSup+' T_CORRESP1, T_AUXILIAIRE' ) ;
                     2 : Q.Sql.Add('Order By '+StSup+' T_CORRESP2, T_AUXILIAIRE' ) ;
                     Else Q.Sql.Add('Order By '+StSup+' T_AUXILIAIRE' ) ;
                     END ;
        Else if CritEdt.GlV.TriePar=0 then Q.SQL.Add(' Order By '+StSup+' T_AUXILIAIRE') Else
             if CritEdt.GlV.TriePar=1 then Q.SQL.Add(' Order By '+StSup+' T_LIBELLE') ;
        End ;
      ChangeSQL(Q) ; Q.Open ;
      END ;
  1 : BEGIN     { Mode de Paiement demandé }
      Q.Close ; Q.SQL.Clear ;
      Q.SQL.Add('Select MP_MODEPAIE , MP_LIBELLE ') ;
      Q.SQL.Add(' From MODEPAIE ') ;
      { Construction de la clause Order By de la SQL }
      If CritEDT.GlV.SQLMODEPAIE<>'' Then Q.SQL.ADD('Where MP_MODEPAIE<>"'+w_w+'" '+CritEDT.GlV.SQLMODEPAIE) ;
      Case CritEdt.GlV.TriePar of
        0 : BEGIN Q.SQL.Add(' Order By MP_MODEPAIE ') ; END ;
        1 : BEGIN Q.SQL.Add(' Order By MP_LIBELLE  ') ; END ;
       end ;
      ChangeSQL(Q) ; Q.Open ;
      END ;
 end ;
END ;

Procedure genereSQLTiersMultiDevise(NE : TNatureEtat ; NET : TNatureEtatTiers ; Var CritEdt : TCritEdt ; Q : TQuery) ;
BEGIN
Q.Close ; Q.SQL.Clear ;
Q.SQL.Add(' Select E_DEVISE as DEV,T_AUXILIAIRE , T_LIBELLE,T_CREDITACCORDE, T_ISPAYEUR as ISPAYEUR, T_PAYEUR AS PAYEUR ') ;
Q.SQL.Add(' From ECRITURE ') ;
Q.SQL.Add(' Left Join TIERS on E_AUXILIAIRE=T_AUXILIAIRE Where ') ;
{ Construction de la clause Where de la SQL }
Q.SQL.Add(' E_AUXILIAIRE<>"" AND E_DEVISE <> "" ') ;
if CritEdt.Joker then Q.SQL.Add(' AND T_AUXILIAIRE like "'+TraduitJoker(CritEdt.Cpt1)+'" ') Else
   BEGIN
   if CritEdt.Cpt1<>'' then Q.SQL.Add(' AND T_AUXILIAIRE>="'+CritEdt.Cpt1+'" ') ;
   if CritEdt.Cpt2<>'' then Q.SQL.Add(' AND T_AUXILIAIRE<="'+CritEdt.Cpt2+'" ') ;
   END ;
if CritEdt.NatureCpt<>'' then Q.SQL.Add(' AND T_NATUREAUXI="'+CritEdt.NatureCpt+'" ')
                         Else Q.SQL.Add(' AND T_NATUREAUXI<>"NCP" ') ;
                         //Else Q.SQL.Add(' AND T_NATUREAUXI<>"SPC" And T_NATUREAUXI<>"PSP" ') ;
{ Rony 26/05/97 , car evite d'avoir un tiers client avec un collectif fournisseur, par exemple}
(* GP le 08/09/99
if CritEdt.SJoker then
   BEGIN
   Q.SQL.Add(' And T_COLLECTIF like "'+TraduitJoker(CritEdt.SCpt1)+'" ' )
   END Else
   BEGIN
   if CritEdt.SCpt1<>'' then Q.SQL.Add(' And T_COLLECTIF>="'+CritEdt.SCpt1+'" ') ;
   if CritEdt.SCpt2<>'' then Q.SQL.Add(' And T_COLLECTIF<="'+CritEdt.SCpt2+'" ') ;
   END ;
*)
   Q.SQL.Add(' AND T_LETTRABLE ="X" ') ;
   Q.SQL.Add(' AND (T_TOTALDEBIT<>0 or T_TOTALCREDIT<>0) ') ;
// BPY le 15/01/2004 => fiche 12403 : ajout de la confidentialité ....
// => old one ...
//      if (OkV2 and (V_PGI.Confidentiel<>'1')) then Q.SQL.Add(' AND T_CONFIDENTIEL<>"1" ') ;
// => pour les future confidantialité !
//      if (V_PGI.Confidentiel = '1') then Q.SQL.Add(' AND T_CONFIDENTIEL<=' + IntToStr(V_PGI.NiveauAccesConf))
//      else Q.SQL.Add(' AND T_CONFIDENTIEL=0');
// => bah ...
      if (V_PGI.Confidentiel <> '1') then Q.SQL.Add(' AND T_CONFIDENTIEL <> "1" ');
// fin BPY
{ Construction de la clause Order By de la SQL }
Q.SQL.Add(' GROUP BY E_DEVISE,T_AUXILIAIRE,T_LIBELLE,T_CREDITACCORDE, T_ISPAYEUR, T_PAYEUR') ;
ChangeSQL(Q) ; Q.Open ;
END ;

Procedure genereSQLSUBTiers(NE : TNatureEtat ; NET : TNatureEtatTiers ; Var CritEdt : TCritEdt ; QEcr : TQuery ; OkTP : Boolean) ;
Var StSupCouv : String ;
BEGIN
{ Construction de la clause Select de la SQL }

QEcr.Close ;
QEcr.SQL.Clear ;
StSupCouv:='' ;
Case CritEdt.Monnaie of
  0 : BEGIN
      StSupCouv:='AND (NOT (E_ETATLETTRAGE="PL" AND ((E_DEBIT+E_CREDIT)=E_COUVERTURE)))'
      END ;
  1 : BEGIN
      StSupCouv:='AND (NOT (E_ETATLETTRAGE="PL" AND ((E_DEBITDEV+E_CREDITDEV)=E_COUVERTUREDEV)))'
      END ;
end ;

Case NE Of
  neGL : BEGIN
          QEcr.SQL.Add('Select E_AUXILIAIRE, E_MODEPAIE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_GENERAL, ') ;
          QEcr.SQL.Add('E_VALIDE, E_REFINTERNE, E_JOURNAL, E_DATEECHEANCE, E_NUMECHE, E_EXERCICE, ') ;
          If OkTP Then QEcr.SQL.Add('E_PIECETP, ') ;
          Case CritEdt.Monnaie of
            0 : BEGIN QEcr.SQL.Add(' E_DEBIT DEBIT, E_CREDIT CREDIT, E_COUVERTURE COUVERTURE ') ; END ;
            1 : BEGIN QEcr.SQL.Add(' E_DEBITDEV DEBIT, E_CREDITDEV CREDIT, E_COUVERTUREDEV COUVERTURE ')   ; END ;
           end ;
          { Tables explorées par la SQL }
          QEcr.SQL.Add(' From ECRITURE') ;
          { Construction de la clause Where de la SQL }
          case CritEdt.GlV.TypePar of
            0 : BEGIN
                If OkTP Then QEcr.SQL.Add('Where E_AUXILIAIRE=:PAYEUR AND E_CONTREPARTIEAUX=:T_AUXILIAIRE')
                        Else QEcr.SQL.Add('Where E_AUXILIAIRE=:T_AUXILIAIRE ') ;
                END ;
            1 : BEGIN
                QEcr.SQL.Add('Where E_MODEPAIE=:MP_MODEPAIE ') ;
                if CritEdt.Joker then QEcr.SQL.Add(' AND E_AUXILIAIRE like "'+TraduitJoker(CritEdt.Cpt1)+'" ') Else
                   BEGIN
                   if CritEdt.Cpt1<>'' then QEcr.SQL.Add(' AND E_AUXILIAIRE>="'+CritEdt.Cpt1+'" ') ;
                   if CritEdt.Cpt2<>'' then QEcr.SQL.Add(' AND E_AUXILIAIRE<="'+CritEdt.Cpt2+'" ') ;
                   END ;
                QEcr.SQL.Add(' And E_AUXILIAIRE<>"" ' ) ;
                END ;
           end ;
          if VH^.ExoV8.Code<>'' then QEcr.SQL.Add(' And E_DATECOMPTABLE>="'+UsDateTime(VH^.ExoV8.Deb)+'" ') ;
          if CritEdt.SJoker then
             BEGIN
             QEcr.SQL.Add(' And E_GENERAL like "'+TraduitJoker(CritEdt.SCpt1)+'" ' )
             END Else
             BEGIN
             if CritEdt.SCpt1<>'' then QEcr.SQL.Add(' And E_GENERAL>="'+CritEdt.SCpt1+'" ') ;
             if CritEdt.SCpt2<>'' then QEcr.SQL.Add(' And E_GENERAL<="'+CritEdt.SCpt2+'" ') ;
             END ;
          QEcr.SQL.Add(TraduitNatureEcr(CritEdt.Qualifpiece, 'E_QUALIFPIECE', true, cbUnchecked(*AvecRevision.State*))) ;
          if CritEdt.Etab<>'' then QEcr.SQL.Add(' And E_ETABLISSEMENT="'+CritEdt.Etab+'"') ;
          if CritEdt.DeviseSelect<>'' then QEcr.SQL.Add(' And E_DEVISE="'+CritEdt.DeviseSelect+'"') ;
          QEcr.SQL.Add(' And E_ECRANOUVEAU<>"CLO" And E_ECRANOUVEAU<>"OAN" ') ;
          QEcr.SQL.Add(' And E_ETATLETTRAGE<>"TL" And E_ETATLETTRAGE<>"RI" ') ;
          QEcr.SQL.Add(' And E_ECHE="X" and E_QUALIFPIECE<>"C" ') ;
          Case CritEdt.GlV.Sens of
            0 : BEGIN
                Case CritEdt.Monnaie of
                  0 : BEGIN QEcr.SQL.Add(' And E_CREDIT<>0 ' ) ;    END ;
                  1 : BEGIN QEcr.SQL.Add(' And E_CREDITDEV<>0 ') ;  END ;
                 end ;
                END ;
            1 : BEGIN
                Case CritEdt.Monnaie of
                  0 : BEGIN QEcr.SQL.Add(' And E_DEBIT<>0 ' ) ;    END ;
                  1 : BEGIN QEcr.SQL.Add(' And E_DEBITDEV<>0 ') ;  END ;
                 end ;
                END ;
           end ;
          If CritEDT.SQLPlus<>'' Then QEcr.SQL.Add(CritEDT.SQLPlus) ;
          If StSupCouv<>'' Then QEcr.SQL.add(StSupCouv) ;
          { Construction de la clause Order By de la SQL }
          case CritEdt.GlV.TypePar of
            0 : QEcr.SQL.Add(' Order By E_AUXILIAIRE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE ') ;
            1 : QEcr.SQL.Add(' Order By E_MODEPAIE,   E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE ') ;
            end ;
          END ;
  neBal:  BEGIN
          QEcr.SQL.Add(' SELECT E_DATEECHEANCE, E_DATECOMPTABLE, E_MODEPAIE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE,') ;
          Case CritEdt.Monnaie of
            0 : BEGIN QEcr.SQL.Add(' E_DEBIT DEBIT, E_CREDIT CREDIT, E_COUVERTURE COUVERTURE ') ; END ;
            1 : BEGIN QEcr.SQL.Add(' E_DEBITDEV DEBIT, E_CREDITDEV CREDIT, E_COUVERTUREDEV COUVERTURE ')   ; END ;
           end ;
          If CritEdt.GlV.DeviseEtPivot Or CritEdt.GlV.Multidevise Then
             BEGIN
             QEcr.SQL.Add(' ,E_DEBIT as DEBITPIVOT, E_CREDIT as CREDITPIVOT, E_COUVERTURE as COUVERTUREPIVOT ') ;
             END ;
          { Tables explorées par la SQL }
          QEcr.SQL.Add(' From ECRITURE ') ;
          { Construction de la clause Where de la SQL }
          case CritEdt.GlV.TypePar of  { 1 : Tiers demandé   ;   1 : Mode de Paiement demandé }
            0 : BEGIN
                If OkTP Then QEcr.SQL.Add('Where E_AUXILIAIRE=:PAYEUR AND E_CONTREPARTIEAUX=:T_AUXILIAIRE')
                        Else QEcr.SQL.Add('Where E_AUXILIAIRE=:T_AUXILIAIRE ') ;
                END ;
            1 : BEGIN
                QEcr.SQL.Add('Where E_MODEPAIE=:MP_MODEPAIE ') ;
                if CritEdt.Joker then QEcr.SQL.Add(' AND E_AUXILIAIRE like "'+TraduitJoker(CritEdt.Cpt1)+'" ') Else
                   BEGIN
                   if CritEdt.Cpt1<>'' then QEcr.SQL.Add(' AND E_AUXILIAIRE>="'+CritEdt.Cpt1+'" ') ;
                   if CritEdt.Cpt2<>'' then QEcr.SQL.Add(' AND E_AUXILIAIRE<="'+CritEdt.Cpt2+'" ') ;
                   END ;
                QEcr.SQL.Add(' AND E_AUXILIAIRE<>"" ') ;
                END ;
           end ;
          if VH^.ExoV8.Code<>'' then QEcr.SQL.Add(' And E_DATECOMPTABLE>="'+UsDateTime(VH^.ExoV8.Deb)+'" ') ;
          if CritEdt.SJoker then
             BEGIN
             QEcr.SQL.Add(' And E_GENERAL like "'+TraduitJoker(CritEdt.SCpt1)+'" ' )
             END Else
             BEGIN
             if CritEdt.SCpt1<>'' then QEcr.SQL.Add(' And E_GENERAL>="'+CritEdt.SCpt1+'" ') ;
             if CritEdt.SCpt2<>'' then QEcr.SQL.Add(' And E_GENERAL<="'+CritEdt.SCpt2+'" ') ;
             END ;
          QEcr.SQL.Add(TraduitNatureEcr(CritEdt.Qualifpiece, 'E_QUALIFPIECE', true, cbUnchecked(*AvecRevision.State*))) ;
          if CritEdt.Etab<>'' then QEcr.SQL.Add(' And E_ETABLISSEMENT="'+CritEdt.Etab+'"') ;
          If CritEdt.GlV.MultiDevise Then
             BEGIN
             QEcr.SQL.Add(' And E_DEVISE=:DEV ') ;
             END Else
             BEGIN
             if CritEdt.DeviseSelect<>'' then QEcr.SQL.Add(' And E_DEVISE="'+CritEdt.DeviseSelect+'"') ;
             END ;
          QEcr.SQL.Add(' And E_ECRANOUVEAU<>"CLO" And E_ECRANOUVEAU<>"OAN" ') ;
          QEcr.SQL.Add(' And E_ETATLETTRAGE<>"TL" ') ;
          QEcr.SQL.Add(' And E_ETATLETTRAGE<>"RI" ') ;
          QEcr.SQL.Add(' And E_ECHE="X" and E_QUALIFPIECE<>"C" ') ;
          Case CritEdt.GlV.Sens of
            0 : BEGIN
                Case CritEdt.Monnaie of
                  0 : BEGIN QEcr.SQL.Add(' And E_CREDIT<>0 ' ) ;    END ;
                  1 : BEGIN QEcr.SQL.Add(' And E_CREDITDEV<>0 ') ;  END ;
                 end ;
                END ;
            1 : BEGIN
                Case CritEdt.Monnaie of
                  0 : BEGIN QEcr.SQL.Add(' And E_DEBIT<>0 ' ) ;    END ;
                  1 : BEGIN QEcr.SQL.Add(' And E_DEBITDEV<>0 ') ;  END ;
                 end ;
                END ;
           end ;
          Case CritEdt.GlV.Sens of
            0 : If CritEDT.GlV.SQLMODEPAIE<>'' Then QEcr.SQL.ADD(CritEDT.GlV.SQLMODEPAIE) ;
            END ;
          If StSupCouv<>'' Then QEcr.SQL.add(StSupCouv) ;
          { Construction de la clause Order By de la SQL }
          QEcr.SQL.Add(' Order By E_DATEECHEANCE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE') ;
          END ;
  END ;
ChangeSQL(QEcr) ; QEcr.Open ;
END ;

procedure CalculDateTiers(NE : TNatureEtat ; NET : TNatureEtatTiers ; Var CritEdt : TCritEdt ;
                          Var TabDate : TTabDate4 ;
                          FP1,FP2,FP3,FP4,FP5,FP6,FP7,FP8 : TMaskEdit) ;
Var i : Integer ;
BEGIN
Case NET Of
  Age :     BEGIN
            if CritEdt.GLV.ChoixEcart=0 then
               BEGIN
               for i:=1 to 4 do
                   TabDate[i]:=CritEdt.Date1-CritEdt.GlV.Ecart*i ;
               END else
               BEGIN
               TabDate[1]:=StrToDate(FP4.Text) ;
               TabDate[2]:=StrToDate(FP3.Text) ;
               TabDate[3]:=StrToDate(FP2.Text) ;
               TabDate[4]:=StrToDate(FP1.Text) ;
               END ;
            END ;
  Ventile : BEGIN
            if CritEdt.GLV.ChoixEcart=0 then
               BEGIN
               Case CritEdt.GlV.TypeGL of
                 0 : BEGIN
                     for i:=1 to 4 do BEGIN TabDate[i]:=CritEdt.Date1+CritEdt.GlV.Ecart*i ; END ;
                     END ;
                 1 : BEGIN
                     for i:=1 to 4 do BEGIN TabDate[i]:=CritEdt.Date1-CritEdt.GlV.Ecart*i ; END ;
                     END ;
                end ;
               END else
               BEGIN
               Case CritEdt.GlV.TypeGL of
                 0 : BEGIN
                     TabDate[1]:=StrToDateTime(FP5.Text) ;
                     TabDate[2]:=StrToDateTime(FP6.Text) ;
                     TabDate[3]:=StrToDateTime(FP7.Text) ;
                     TabDate[4]:=StrToDateTime(FP8.Text) ;
                     END ;
                 1 : BEGIN
                     TabDate[1]:=StrToDateTime(FP4.Text) ;
                     TabDate[2]:=StrToDateTime(FP3.Text) ;
                     TabDate[3]:=StrToDateTime(FP2.Text) ;
                     TabDate[4]:=StrToDateTime(FP1.Text) ;
                     END ;
                end ;
               END ;
            END ;
  END ;
END ;

Function ValidPeriodeTiers(FP1,FP2,FP3,FP4,FP5,FP6,FP7,FP8 : TMaskEdit) : Boolean ;
begin
if      ( StrToDate(FP1.Text) < StrToDate(FP5.Text) )
    And ( StrToDate(FP2.Text) < StrToDate(FP6.Text) )
    And ( StrToDate(FP3.Text) < StrToDate(FP7.Text) )
    And ( StrToDate(FP4.Text) < StrToDate(FP8.Text) )
    And ( StrToDate(FP1.Text) < StrToDate(FP8.Text) )
  then Result:=True
  else Result:=False ;
end;

procedure ChargeRecapMPTiers (Var L : TStringList; MultiCombo : THMultiValcomboBox) ;
{Création de la liste des modes de paiement }
Var LModPaie         : TModPaie ;
    ChaineMP, CodeMP : String ;
BEGIN
L:=TStringList.Create ; L.Sorted:=TRUE ; L.Clear ;
ChaineMP:=MultiCombo.Text ;
CodeMP:=ReadTokenSt(ChaineMP) ;
While (CodeMP<>'') do
  BEGIN
  LModPaie:=TModPaie.Create ;
  LModPaie.Montant:=0 ;
  L.AddObject(CodeMP,LModPaie) ;
  CodeMP:=ReadTokenSt(ChaineMP) ;
  END ;
END ;

function MDPRetenuTiers(Var L : TStringList ; St : string) : Boolean ;
{ Consultation pour un mode de paiement si dans Liste True sinon False }
Var i : integer ;
BEGIN
Result:=False ;
i:=0 ;
While i<=L.Count-1 do
  BEGIN
  if Copy(St,1,Length(L[i]))=Copy(L[i],1,Length(L[i])) then
     BEGIN
     Result:=True ;
     Exit ;
     END else
     BEGIN
     Inc(i) ;
     END;
  END ;
END ;

procedure PeriodiciteChangeTiers(NE : TNatureEtat ; NET : TNatureEtatTiers ; Iperiodicite : Integer ;
                                 FP8,FP1 : TMaskEdit ; Retard : Boolean ; Var TabD : TTabDate8) ;
Var Choix,i : Integer ;
    a,m,j, JMax         : Word ;
    DAT : TDATETIME ;
begin
//Choix:=FPeriodicite.ItemIndex ;
Choix:=IPeriodicite ;
If Choix<=-1 Then Choix:=0 ;
If Retard Then
   BEGIN
   TabD[8]:=StrToDate(FP8.Text) ;       { Calcul à partir de la date d'arrêtée }
   Case IPeriodicite of
     0,1,2,3,4,5 : for i:=4 downto 1 Do { Pour chaque Fourchette de dates( en partant de la derniére) , en mensuel }
                       BEGIN
                       DAT:=PlusMois(TabD[i+4],-(Choix+1)) ;
                       DecodeDate(DAT,a,m,j) ;
                       JMax:=StrToInt(FormatDateTime('d',FinDeMois(EncodeDate(a,m,1)))) ;
                       TabD[i]:=PlusMois(TabD[i+4],-(Choix+1))+(JMax-J)+1 ;
                       if i>1 then TabD[i+3]:=TabD[i]-1 ;
                       END ;
     6           : for i:=4 downto 1 Do { Pour chaque Fourchette de dates( en partant de la derniére) , en Quinzaine }
                       BEGIN
                       DAT:=TabD[i+4] ;
                       DecodeDate(DAT,a,m,j) ;
                       JMax:=StrToInt(FormatDateTime('d',FinDeMois(EncodeDate(a,m,1)))) ;
                       If J<=15 then
                          BEGIN         { Date départ = (Date d'arrivée - 15 jours) + 1 jour, si date avant le 15 du mois }
                          TabD[i]:=TabD[i+4]-(15  )+1 ;
                          END Else
                          BEGIN         { Date départ = (Date d'arrivée - (Nb jours Max du mois - 15 jours) ) + 1 jour, si date aprés le 15 du mois }
                          TabD[i]:=TabD[i+4]-(JMax-15)+1 ;
                          END ;
                       if i>1 then TabD[i+3]:=TabD[i]-1 ;
                       END ;
     7           : for i:=4 downto 1 Do { Pour chaque Fourchette de dates( en partant de la derniére) , en Hebdo }
                       BEGIN            { Date départ = (Date d'arrivée - (7 jours+ 1 jours) )  }
                       TabD[i]:=TabD[i+4]-(7)+1 ;
                       if i>1 then TabD[i+3]:=TabD[i]-1 ;
                       END ;
     End ;
   END Else
   BEGIN
   TabD[1]:=StrToDate(FP1.Text) ;   { Calcul à partir de la date d'arrêtée }
   Case IPeriodicite of   { 0..5 : Période Mensuel ; 6 : Quinzaine ; 7 ; Hebdomadaire }
     0,1,2,3,4,5 : BEGIN
                   For i:=5 to 8 Do { Pour chaque Fourchette de dates, en mensuel }
                       BEGIN        { Date d'arrivée = (Date départ + nb Mois) - 1 jour }
                       If i=5 Then
                         BEGIN
                         DAT:=PlusMois(TabD[i-4], (Choix)) ;
                         TabD[i]:=FinDeMois(DAT) ;
                         END Else TabD[i]:=PlusMois(TabD[i-4], (Choix+1))-1 ;
                       if i<8 then TabD[i-3]:=TabD[i]+1 ;
                       END ;
                   END ;
     6           : BEGIN
                   for i:=5 to 8 Do { Pour chaque Fourchette de dates, en Quinzaine }
                       BEGIN
                       DecodeDate(TabD[i-4],a,m,j) ;
                       JMax:=StrToInt(FormatDateTime('d',FinDeMois(EncodeDate(a,m,1)))) ;
                       If J<15 then
                          BEGIN    { Date d'arrivée = (Date départ + 15 jours) - 1 jour, si date avant le 15 du mois }
                          TabD[i]:=TabD[i-4]+(15-1) ;
                          END ELse
                          BEGIN    { Date d'arrivée = (Date départ + (Nb jours Max du mois - 15 jours) ) - 1 jour, si date aprés le 15 du mois }
                          TabD[i]:=TabD[i-4]+(JMax-15)-1 ;
                          END ;
                       if i<8 then TabD[i-3]:=TabD[i]+1 ;
                       END ;
                   END ;
     7           : BEGIN
                   for i:=5 to 8 Do { Pour chaque Fourchette de dates, en Hebdo }
                       BEGIN        { Date d'arrivée = (Date départ + (7 jours - 1 jour)  }
                       TabD[i]:=TabD[i-4]+(7-1) ;
                       if i<8 then TabD[i-3]:=TabD[i]+1 ;
                       END ;
                   END ;
     End ;
   END ;
end;

Procedure ChangeColl(Nature : String ; C1, C2 : THCpteEdit) ;
BEGIN
if Nature='CLI' then begin C1.ZoomTable:=tzGCollClient ; C2.ZoomTable:=C1.ZoomTable ; end else
if Nature='FOU' then begin C1.ZoomTable:=tzGCollFourn ; C2.ZoomTable:=C1.ZoomTable ; end else
if Nature='SAL' then begin C1.ZoomTable:=tzGCollSalarie ; C2.ZoomTable:=C1.ZoomTable ; end else
if Nature='DIV' then begin C1.ZoomTable:=tzGCollDivers ; C2.ZoomTable:=C1.ZoomTable ; end else
if Nature='AUC' then begin C1.ZoomTable:=tzGCollToutCredit ; C2.ZoomTable:=C1.ZoomTable ; end else
if Nature='AUD' then begin C1.ZoomTable:=tzGCollToutDebit ; C2.ZoomTable:=C1.ZoomTable ; end else
if Nature='' then begin C1.ZoomTable:=tzGCollectif ; C2.ZoomTable:=C1.ZoomTable ; end else
END;


Procedure RetailleBandesBudget(FF : TForm ; CritEdt : TCritEdt ; LibRev : String ; BalOuEca : Boolean ; BD, BDLibre, BDFin : TQRBAND) ;
Const TopIni = 2 ; TopRev = 16 ; TopRea = 30 ; TopEca = 44 ; TopQIni = 30 ; TopQRev = 44 ; TopPlus = 58 ;
Var I, Min, Max        : Byte ;
    Largeur            : Integer ;
    LaBudNat           : String ;
    Balance            : Boolean ;
    C1, C2, C3, C4,
    C5, C6, C7, C8, C9, C10 : TComponent ;
BEGIN
Balance:=(BalOuEca=True);
If Balance then
   BEGIN
   Min:=CritEdt.BalBud.FormatPrint.ColMin ;
   Max:=CritEdt.BalBud.FormatPrint.ColMax-1 ;
   LaBudNat:=CritEdt.BalBud.NatureBud ;
   { Bande Détail }
   TQrLabel(FF.FindComponent('IniGen')).Visible:=(LaBudNat='') or (LaBudNat='INI') ;
   TQrLabel(FF.FindComponent('IniGenTotal')).Visible:=(LaBudNat='') or (LaBudNat='INI') ;

   TQrLabel(FF.FindComponent('RevGen')).Visible:=(LaBudNat='') or (Copy(LaBudNat,1,2)='DM') ;
   TQrLabel(FF.FindComponent('RevGenTotal')).Visible:=(LaBudNat='') or (Copy(LaBudNat,1,2)='DM') ;

   TQrLabel(FF.FindComponent('ReaGen')).Visible:=CritEdt.BalBud.Realise ;
   TQrLabel(FF.FindComponent('ReaGenTotal')).Visible:=CritEdt.BalBud.Realise ;
   TQrLabel(FF.FindComponent('EcaGen')).Visible:=CritEdt.BalBud.Realise ;
   TQrLabel(FF.FindComponent('EcaGenTotal')).Visible:=CritEdt.BalBud.Realise ;

   { Bande Libre ou Rupture }
   TQrLabel(FF.FindComponent('IniLibre')).Visible:=(LaBudNat='') or (LaBudNat='INI') ;
   TQrLabel(FF.FindComponent('IniLibreTotal')).Visible:=(LaBudNat='') or (LaBudNat='INI') ;

   TQrLabel(FF.FindComponent('RevLibre')).Visible:=(LaBudNat='') or (Copy(LaBudNat,1,2)='DM') ;
   TQrLabel(FF.FindComponent('RevLibreTotal')).Visible:=(LaBudNat='') or (Copy(LaBudNat,1,2)='DM') ;

   TQrLabel(FF.FindComponent('ReaLibre')).Visible:=CritEdt.BalBud.Realise ;
   TQrLabel(FF.FindComponent('ReaLibreTotal')).Visible:=CritEdt.BalBud.Realise ;
   TQrLabel(FF.FindComponent('EcaLibre')).Visible:=CritEdt.BalBud.Realise ;
   TQrLabel(FF.FindComponent('EcaLibreTotal')).Visible:=CritEdt.BalBud.Realise ;

   TQrLabel(FF.FindComponent('RevGen')).Caption:=LibRev ;
   TQrLabel(FF.FindComponent('RevLibre')).Caption:=LibRev ;
   { Total Edition }
   TQrLabel(FF.FindComponent('ReaTotal')).Visible:=CritEdt.BalBud.Realise ;
   TQrLabel(FF.FindComponent('TotTotalR')).Visible:=CritEdt.BalBud.Realise ;
   TQrLabel(FF.FindComponent('EcaTotal')).Visible:=CritEdt.BalBud.Realise ;
   TQrLabel(FF.FindComponent('TotTotalE')).Visible:=CritEdt.BalBud.Realise ;

   IF LaBudNat='INI' then  { Nature du budget à INI }
      BEGIN
      For i:=Min to Max do
          BEGIN
          C1:=FF.FindComponent('IniGenCum'+InttoStr(i))   ; TQRLabel(C1).Top:=TopIni ; TQRLabel(C1).Visible:=True ;
          C2:=FF.FindComponent('RevGenCum'+InttoStr(i))   ; TQRLabel(C2).Visible:=False ;
          C3:=FF.FindComponent('ReaGenCum'+InttoStr(i))   ; TQRLabel(C3).Visible:=CritEdt.BalBud.Realise ;
          C4:=FF.FindComponent('EcaGenCum'+InttoStr(i))   ; TQRLabel(C4).Visible:=CritEdt.BalBud.Realise ;
          { Bande nature libre }
          C5:=FF.FindComponent('IniLibre'+InttoStr(i))   ; TQRLabel(C5).Top:=TQRLabel(C1).Top ; TQRLabel(C5).Visible:=TQRLabel(C1).Visible ;
          C6:=FF.FindComponent('RevLibre'+InttoStr(i))   ; TQRLabel(C6).Visible:=TQRLabel(C2).Visible ;
          C7:=FF.FindComponent('ReaLibre'+InttoStr(i))   ; TQRLabel(C7).Visible:=TQRLabel(C3).Visible ;
          C8:=FF.FindComponent('EcaLibre'+InttoStr(i))   ; TQRLabel(C8).Visible:=TQRLabel(C4).Visible ;
          if CritEdt.BalBud.Realise then
             BEGIN
             TQRLabel(C3).Top:=TopRev ;
             TQRLabel(C4).Top:=TopRea ;
             END Else
             BEGIN
             TQRLabel(C3).Top:=TopRea ;
             TQRLabel(C4).Top:=TopEca ;
             END ;
          TQRLabel(C7).Top:=TQRLabel(C3).Top ;
          TQRLabel(C8).Top:=TQRLabel(C4).Top ;
          END ;
      TQrLabel(FF.FindComponent('IniGenTotal')).Top:=TopIni ; TQrLabel(FF.FindComponent('IniGen')).Top:=TopIni ;
      {Nature libre}
      TQrLabel(FF.FindComponent('IniLibreTotal')).Top:=TopIni ;
      TQrLabel(FF.FindComponent('IniLibre')).Top:=TopIni ;
      if CritEdt.BalBud.Realise then
         BEGIN
         TQrLabel(FF.FindComponent('ReaGenTotal')).Top:=TopRev ;
         TQrLabel(FF.FindComponent('ReaGen')).Top:=TopRev ;
         TQrLabel(FF.FindComponent('EcaGenTotal')).Top:=TopRea ;
         TQrLabel(FF.FindComponent('EcaGen')).Top:=TopRea ;
         Largeur:=TQrLabel(FF.FindComponent('EcaGen')).Top+TQrLabel(FF.FindComponent('EcaGen')).Height+1 ;
         END Else
         BEGIN
         TQrLabel(FF.FindComponent('ReaGenTotal')).Top:=TopRea ;
         TQrLabel(FF.FindComponent('ReaGen')).Top:=TopRea ;
         TQrLabel(FF.FindComponent('EcaGenTotal')).Top:=TopEca ;
         TQrLabel(FF.FindComponent('EcaGen')).Top:=TopEca ;
         Largeur:=TQrLabel(FF.FindComponent('IniGen')).Top+TQrLabel(FF.FindComponent('IniGen')).Height+1 ;
         END ;
      BD.Height:=Largeur ;
      {Nature libre}
      TQrLabel(FF.FindComponent('ReaLibreTotal')).Top:=TQrLabel(FF.FindComponent('ReaGenTotal')).Top ;
      TQrLabel(FF.FindComponent('ReaLibre')).Top:=TQrLabel(FF.FindComponent('ReaGen')).Top ;
      TQrLabel(FF.FindComponent('EcaLibreTotal')).Top:=TQrLabel(FF.FindComponent('EcaGenTotal')).Top ;
      TQrLabel(FF.FindComponent('EcaLibre')).Top:=TQrLabel(FF.FindComponent('EcaGen')).Top ;
      END Else
   If Copy(LaBudNat,1,2)='DM' then { Nature du budget à DM X }
      BEGIN
      For i:=Min to Max do
          BEGIN
          C1:=FF.FindComponent('IniGenCum'+InttoStr(i)); TQRLabel(C1).Visible:=False ;
          C2:=FF.FindComponent('RevGenCum'+InttoStr(i)); TQRLabel(C2).Top:=TopIni ; TQRLabel(C2).Visible:=True ;
          C3:=FF.FindComponent('ReaGenCum'+InttoStr(i)); TQRLabel(C3).Visible:=CritEdt.BalBud.Realise ;
          C4:=FF.FindComponent('EcaGenCum'+InttoStr(i)); TQRLabel(C4).Visible:=CritEdt.BalBud.Realise ;
          { Nature Libe}
          C5:=FF.FindComponent('IniLibre'+InttoStr(i)) ; TQRLabel(C5).Visible:=TQRLabel(C1).Visible ;
          C6:=FF.FindComponent('RevLibre'+InttoStr(i)) ; TQRLabel(C6).Top:=TQRLabel(C2).Top ; TQRLabel(C6).Visible:=TQRLabel(C2).Visible ;
          C7:=FF.FindComponent('ReaLibre'+InttoStr(i)) ; TQRLabel(C7).Visible:=TQRLabel(C3).Visible ;
          C8:=FF.FindComponent('EcaLibre'+InttoStr(i)) ; TQRLabel(C8).Visible:=TQRLabel(C4).Visible ;
          if CritEdt.BalBud.Realise then
             BEGIN
             TQRLabel(C3).Top:=TopRev ;
             TQRLabel(C4).Top:=TopRea ;
             END Else
             BEGIN
             TQRLabel(C3).Top:=TopRea ;
             TQRLabel(C4).Top:=TopEca ;
             END ;
          TQRLabel(C7).Top:=TQRLabel(C3).Top ;
          TQRLabel(C8).Top:=TQRLabel(C4).Top ;
          END ;
      TQrLabel(FF.FindComponent('RevGenTotal')).Top:=TopIni ;
      TQrLabel(FF.FindComponent('RevGen')).Top:=TopIni ;
      TQrLabel(FF.FindComponent('RevGen')).Caption:=LibRev+' '+LaBudNat ;
      { Nature Libre }
      TQrLabel(FF.FindComponent('RevLibreTotal')).Top:=TopIni ;
      TQrLabel(FF.FindComponent('RevLibre')).Top:=TopIni ;
      TQrLabel(FF.FindComponent('RevLibre')).Caption:=LibRev+' '+LaBudNat ;

      if CritEdt.BalBud.Realise then
         BEGIN
         TQrLabel(FF.FindComponent('ReaGenTotal')).Top:=TopRev ;
         TQrLabel(FF.FindComponent('ReaGen')).Top:=TopRev ;
         TQrLabel(FF.FindComponent('EcaGenTotal')).Top:=TopRea ;
         TQrLabel(FF.FindComponent('EcaGen')).Top:=TopRea ;
         Largeur:=TQrLabel(FF.FindComponent('EcaGen')).Top+TQrLabel(FF.FindComponent('EcaGen')).Height+1 ;
         END Else
         BEGIN
         Largeur:=TQrLabel(FF.FindComponent('RevGen')).Top+TQrLabel(FF.FindComponent('RevGen')).Height+1 ;
         END ;
      BD.Height:=Largeur ;
      { Nature Libre }
      TQrLabel(FF.FindComponent('ReaLibreTotal')).Top:=TQrLabel(FF.FindComponent('ReaGenTotal')).Top ;
      TQrLabel(FF.FindComponent('ReaLibre')).Top:=TQrLabel(FF.FindComponent('ReaGen')).Top ;

      TQrLabel(FF.FindComponent('EcaLibreTotal')).Top:=TQrLabel(FF.FindComponent('EcaGenTotal')).Top ;
      TQrLabel(FF.FindComponent('EcaLibre')).Top:=TQrLabel(FF.FindComponent('EcaGen')).Top ;
      END Else
      BEGIN      { Nature du budget à TOUS }
      For i:=Min to Max do
          BEGIN
          C1:=FF.FindComponent('IniGenCum'+InttoStr(i)) ; TQRLabel(C1).Visible:=True ; TQRLabel(C1).Top:=TopIni ;
          C2:=FF.FindComponent('RevGenCum'+InttoStr(i)) ; TQRLabel(C2).Visible:=True ; TQRLabel(C2).Top:=TopRev ;
          C3:=FF.FindComponent('ReaGenCum'+InttoStr(i)) ; TQRLabel(C3).Visible:=CritEdt.BalBud.Realise ;
          C4:=FF.FindComponent('EcaGenCum'+InttoStr(i)) ; TQRLabel(C4).Visible:=CritEdt.BalBud.Realise ;
          TQRLabel(C3).Top:=TopRea ; TQRLabel(C4).Top:=TopEca ;
          { Nature Libre }
          C5:=FF.FindComponent('IniLibre'+InttoStr(i)) ; TQRLabel(C5).Visible:=TQRLabel(C1).Visible ; TQRLabel(C5).Top:=TQRLabel(C1).Top ;
          C6:=FF.FindComponent('RevLibre'+InttoStr(i)) ; TQRLabel(C6).Visible:=TQRLabel(C2).Visible ; TQRLabel(C6).Top:=TQRLabel(C2).Top ;
          C7:=FF.FindComponent('ReaLibre'+InttoStr(i)) ; TQRLabel(C7).Visible:=TQRLabel(C3).Visible ;
          C8:=FF.FindComponent('EcaLibre'+InttoStr(i)) ; TQRLabel(C8).Visible:=TQRLabel(C4).Visible ;
          TQRLabel(C7).Top:=TQRLabel(C3).Top ; TQRLabel(C8).Top:=TQRLabel(C4).Top ;
          END ;
      TQrLabel(FF.FindComponent('IniGenTotal')).Top:=TopIni ;
      TQrLabel(FF.FindComponent('IniGen')).Top:=TopIni ;
      TQrLabel(FF.FindComponent('RevGenTotal')).Top:=TopRev ;
      TQrLabel(FF.FindComponent('RevGen')).Top:=TopRev ;
      TQrLabel(FF.FindComponent('ReaGenTotal')).Top:=TopRea ;
      TQrLabel(FF.FindComponent('ReaGen')).Top:=TopRea ;
      TQrLabel(FF.FindComponent('EcaGenTotal')).Top:=TopEca ;
      TQrLabel(FF.FindComponent('EcaGen')).Top:=TopEca ;

      if CritEdt.BalBud.Realise then Largeur:=TopEca+TQrLabel(FF.FindComponent('EcaGen')).Height+1
                                Else Largeur:=TopRev+TQrLabel(FF.FindComponent('RevGen')).Height+1 ;
      BD.Height:=Largeur ;
      { Nature Libre }
      TQrLabel(FF.FindComponent('IniLibreTotal')).Top:=TopIni ;
      TQrLabel(FF.FindComponent('IniLibre')).Top:=TopIni ;
      TQrLabel(FF.FindComponent('RevLibreTotal')).Top:=TopRev ;
      TQrLabel(FF.FindComponent('RevLibre')).Top:=TopRev ;
      TQrLabel(FF.FindComponent('ReaLibreTotal')).Top:=TopRea ;
      TQrLabel(FF.FindComponent('ReaLibre')).Top:=TopRea ;
      TQrLabel(FF.FindComponent('EcaLibreTotal')).Top:=TopEca ;
      TQrLabel(FF.FindComponent('EcaLibre')).Top:=TopEca ;
      END ;
   BDLibre.Height:=BD.Height ;
   {Affichage ou pas du Total général en réalisé et ecart}
   For i:=Min to Max do
       BEGIN
       C1:=FF.FindComponent('ReaTotal'+InttoStr(i)) ; TQRLabel(C1).Visible:=CritEdt.BalBud.Realise ;
       C2:=FF.FindComponent('EcaTotal'+InttoStr(i)) ; TQRLabel(C2).Visible:=CritEdt.BalBud.Realise ;
       END ;
   if CritEdt.BalBud.Realise then Largeur:=TQrLabel(FF.FindComponent('EcaTotal')).Top+TQrLabel(FF.FindComponent('EcaTotal')).Height+5
                             Else Largeur:=TQrLabel(FF.FindComponent('BudTotal')).Top+TQrLabel(FF.FindComponent('BudTotal')).Height+5 ;
   BDFin.Height:=Largeur ;
   END Else
   BEGIN
{ - - - - - ECARTS BUDGETAIRES - - - - - - }
(* *)
//   Min:=CritEdt.BalBud.FormatPrint.ColMin ;
//   Max:=CritEdt.BalBud.FormatPrint.ColMax-1 ;
   LaBudNat:=CritEdt.BalBud.NatureBud ;
   TQrLabel(FF.FindComponent('RevGen')).Caption:=LibRev ;

   TQrLabel(FF.FindComponent('IniGen')).Visible:=(LaBudNat='') or (LaBudNat='INI') ;
   TQrLabel(FF.FindComponent('RevGen')).Visible:=(LaBudNat='') or (Copy(LaBudNat,1,2)='DM') ;
   TQrLabel(FF.FindComponent('QIniGen')).Visible:=((LaBudNat='') or (LaBudNat='INI')) and CritEdt.BalBud.Qte ;
   TQrLabel(FF.FindComponent('QRevGen')).Visible:=((LaBudNat='') or (Copy(LaBudNat,1,2)='DM')) and CritEdt.BalBud.Qte ;

   TQrLabel(FF.FindComponent('IniGen')).Font.Style:=TQrLabel(FF.FindComponent('IniGen')).Font.Style-[fsItalic] ;
   TQrLabel(FF.FindComponent('RevGen')).Font.Style:=TQrLabel(FF.FindComponent('RevGen')).Font.Style-[fsItalic] ;
   TQrLabel(FF.FindComponent('IRGen')).Visible:=CritEdt.BalBud.TotIniRev ;

   {Nature Libre ou Rupture }
   TQrLabel(FF.FindComponent('RevLibre')).Caption:=LibRev ;

   TQrLabel(FF.FindComponent('IniLibre')).Visible:=(LaBudNat='') or (LaBudNat='INI') ;
   TQrLabel(FF.FindComponent('RevLibre')).Visible:=(LaBudNat='') or (Copy(LaBudNat,1,2)='DM') ;
   TQrLabel(FF.FindComponent('QIniLibre')).Visible:=((LaBudNat='') or (LaBudNat='INI')) and CritEdt.BalBud.Qte ;
   TQrLabel(FF.FindComponent('QRevLibre')).Visible:=((LaBudNat='') or (Copy(LaBudNat,1,2)='DM')) and CritEdt.BalBud.Qte ;

   TQrLabel(FF.FindComponent('IniLibre')).Font.Style:=TQrLabel(FF.FindComponent('IniGen')).Font.Style-[fsItalic] ;
   TQrLabel(FF.FindComponent('RevLibre')).Font.Style:=TQrLabel(FF.FindComponent('RevGen')).Font.Style-[fsItalic] ;
   TQrLabel(FF.FindComponent('IRLibre')).Visible:=CritEdt.BalBud.TotIniRev ;
   IF LaBudNat='INI' then  { Nature du budget à INI }
      BEGIN
      For i:=1 to 10 do
          BEGIN
          C1:=FF.FindComponent('IniGenCum'+InttoStr(i))  ; TQRLabel(C1).Top:=TopIni ; TQRLabel(C1).Visible:=True ;
          C2:=FF.FindComponent('RevGenCum'+InttoStr(i))  ; TQRLabel(C2).Visible:=False ;
          C3:=FF.FindComponent('QIniGenCum'+InttoStr(i)) ; TQRLabel(C3).Visible:=CritEdt.BalBud.Qte ;
          TQRLabel(C1).Font.Style:=TQRLabel(C1).Font.Style-[fsItalic] ;
          if CritEdt.BalBud.Qte then TQRLabel(C3).Top:=TopRev
                                Else TQRLabel(C3).Top:=TopQIni ;
          C4:=FF.FindComponent('QRevGenCum'+InttoStr(i))   ; TQRLabel(C4).Visible:=False ;
          TQRLabel(C4).Top:=TopQRev ;
          { Nature Libre ou Rupture }
          C5:=FF.FindComponent('IniLibre'+InttoStr(i))  ; TQRLabel(C5).Top:=TopIni ; TQRLabel(C1).Visible:=True ;
          C6:=FF.FindComponent('RevLibre'+InttoStr(i))  ; TQRLabel(C6).Visible:=False ;
          C7:=FF.FindComponent('QIniLibre'+InttoStr(i)) ; TQRLabel(C7).Visible:=CritEdt.BalBud.Qte ;
          TQRLabel(C5).Font.Style:=TQRLabel(C5).Font.Style-[fsItalic] ;
          if CritEdt.BalBud.Qte then TQRLabel(C7).Top:=TopRev
                                Else TQRLabel(C7).Top:=TopQIni ;
          C8:=FF.FindComponent('QRevLibre'+InttoStr(i))   ; TQRLabel(C8).Visible:=False ;
          TQRLabel(C8).Top:=TopQRev ;
          END ;
      TQrLabel(FF.FindComponent('IniGen')).Top:=TopIni ;
      {Nature Libre }
      TQrLabel(FF.FindComponent('IniLibre')).Top:=TopIni ;
      if CritEdt.BalBud.Qte then
         BEGIN
         TQrLabel(FF.FindComponent('QIniGen')).Top:=TopRev ;
         Largeur:=TQrLabel(FF.FindComponent('QIniGen')).Top+TQrLabel(FF.FindComponent('QIniGen')).Height+1 ;
         END Else
         BEGIN
         TQrLabel(FF.FindComponent('QIniGen')).Top:=TopQIni ;
         Largeur:=TQrLabel(FF.FindComponent('IniGen')).Top+TQrLabel(FF.FindComponent('IniGen')).Height+1 ;
         END ;
      {Nature Libre }
      TQrLabel(FF.FindComponent('QIniLibre')).Top:=TQrLabel(FF.FindComponent('QIniGen')).Top ;
      END Else
   If Copy(LaBudNat,1,2)='DM' then { Nature du budget à DM X }
      BEGIN
      For i:=1 to 10 do
          BEGIN
          C1:=FF.FindComponent('IniGenCum'+InttoStr(i))   ; TQRLabel(C1).Visible:=False ;
          C2:=FF.FindComponent('RevGenCum'+InttoStr(i))   ; TQRLabel(C2).Top:=TopIni ; TQRLabel(C2).Visible:=True ;
          C3:=FF.FindComponent('QIniGenCum'+InttoStr(i))   ; TQRLabel(C3).Visible:=False ;
          TQRLabel(C2).Font.Style:=TQRLabel(C2).Font.Style-[fsItalic] ;
          C4:=FF.FindComponent('QRevGenCum'+InttoStr(i))   ; TQRLabel(C4).Visible:=CritEdt.BalBud.Qte ;
          if CritEdt.BalBud.Qte then TQRLabel(C4).Top:=TopRev
                                Else TQRLabel(C4).Top:=TopQRev ;
          { Nature Libre}
          C5:=FF.FindComponent('IniGenCum'+InttoStr(i))   ; TQRLabel(C5).Visible:=False ;
          C6:=FF.FindComponent('RevGenCum'+InttoStr(i))   ; TQRLabel(C6).Top:=TopIni ; TQRLabel(C2).Visible:=True ;
          C7:=FF.FindComponent('QIniGenCum'+InttoStr(i))   ; TQRLabel(C7).Visible:=False ;
          TQRLabel(C6).Font.Style:=TQRLabel(C6).Font.Style-[fsItalic] ;
          C8:=FF.FindComponent('QRevGenCum'+InttoStr(i))   ; TQRLabel(C8).Visible:=CritEdt.BalBud.Qte ;
          if CritEdt.BalBud.Qte then TQRLabel(C8).Top:=TopRev
                                Else TQRLabel(C8).Top:=TopQRev ;
          END ;
      TQrLabel(FF.FindComponent('RevGen')).Caption:=LibRev+' '+LaBudNat ;
      TQrLabel(FF.FindComponent('RevGen')).Top:=TopIni ;
      { Nature Libre ou Rupture }
      TQrLabel(FF.FindComponent('RevLibre')).Caption:=LibRev+' '+LaBudNat ;
      TQrLabel(FF.FindComponent('RevLibre')).Top:=TopIni ;
      if CritEdt.BalBud.Qte then
         BEGIN
         TQrLabel(FF.FindComponent('QRevGen')).Top:=TopRev ;
         Largeur:=TQrLabel(FF.FindComponent('QRevGen')).Top+TQrLabel(FF.FindComponent('QRevGen')).Height+1 ;
         END Else
         BEGIN
         TQrLabel(FF.FindComponent('QRevGen')).Top:=TopQRev ;
         Largeur:=TQrLabel(FF.FindComponent('RevGen')).Top+TQrLabel(FF.FindComponent('RevGen')).Height+1 ;
         END ;
      { Nature Libre ou Rupture }
      TQrLabel(FF.FindComponent('QRevLibre')).Top:=TQrLabel(FF.FindComponent('QRevGen')).Top ;
      END Else
      BEGIN      { Nature du budget à TOUS }
      For i:=1 to 10 do
          BEGIN
          C1:=FF.FindComponent('IniGenCum'+InttoStr(i))  ; TQRLabel(C1).Visible:=True ; TQRLabel(C1).Top:=TopIni ;
          C2:=FF.FindComponent('RevGenCum'+InttoStr(i))  ; TQRLabel(C2).Visible:=True ; TQRLabel(C2).Top:=TopRev ;
          C3:=FF.FindComponent('QIniGenCum'+InttoStr(i)) ; TQRLabel(C3).Visible:=CritEdt.BalBud.Qte ;
          C4:=FF.FindComponent('QRevGenCum'+InttoStr(i)) ; TQRLabel(C4).Visible:=CritEdt.BalBud.Qte ;
          C5:=FF.FindComponent('IRGenCum'+InttoStr(i))   ; TQRLabel(C5).Visible:=CritEdt.BalBud.TotIniRev ;
          If CritEdt.BalBud.TotIniRev then
             BEGIN
             TQRLabel(C5).Top:=TopQIni ;
             TQRLabel(C3).Top:=TopQRev ;
             TQRLabel(C4).Top:=TopPlus ;
             TQRLabel(C1).Font.Style:=TQRLabel(C1).Font.Style+[fsItalic] ;
             TQRLabel(C2).Font.Style:=TQRLabel(C2).Font.Style+[fsItalic] ;
             END Else
             BEGIN
             TQRLabel(C3).Top:=TopQIni ;
             TQRLabel(C4).Top:=TopQRev ;
             TQRLabel(C1).Font.Style:=TQRLabel(C1).Font.Style-[fsItalic] ;
             TQRLabel(C2).Font.Style:=TQRLabel(C2).Font.Style-[fsItalic] ;
             END ;
          { Nature Libre ou Rupture }
          C6:=FF.FindComponent('IniLibre'+InttoStr(i))  ; TQRLabel(C6).Visible:=True ; TQRLabel(C1).Top:=TopIni ;
          C7:=FF.FindComponent('RevLibre'+InttoStr(i))  ; TQRLabel(C7).Visible:=True ; TQRLabel(C2).Top:=TopRev ;
          C8:=FF.FindComponent('QIniLibre'+InttoStr(i)) ; TQRLabel(C8).Visible:=CritEdt.BalBud.Qte ;
          C9:=FF.FindComponent('QRevLibre'+InttoStr(i)) ; TQRLabel(C9).Visible:=CritEdt.BalBud.Qte ;
          C10:=FF.FindComponent('IRLibre'+InttoStr(i))   ; TQRLabel(C10).Visible:=CritEdt.BalBud.TotIniRev ;
          If CritEdt.BalBud.TotIniRev then
             BEGIN
             TQRLabel(C10).Top:=TQRLabel(C5).Top ;
             TQRLabel(C8).Top:=TQRLabel(C3).Top ;
             TQRLabel(C9).Top:=TQRLabel(C4).Top ;
             TQRLabel(C6).Font.Style:=TQRLabel(C1).Font.Style ;
             TQRLabel(C7).Font.Style:=TQRLabel(C2).Font.Style ;
             END Else
             BEGIN
             TQRLabel(C8).Top:=TQRLabel(C3).Top ;
             TQRLabel(C9).Top:=TQRLabel(C4).Top ;
             TQRLabel(C6).Font.Style:=TQRLabel(C1).Font.Style ;
             TQRLabel(C7).Font.Style:=TQRLabel(C2).Font.Style ;
             END ;
          END ;
      TQrLabel(FF.FindComponent('IniGen')).Top:=TopIni ;
      TQrLabel(FF.FindComponent('RevGen')).Top:=TopRev ;
      { Nature Libre ou Rupture }
      TQrLabel(FF.FindComponent('IniLibre')).Top:=TopIni ;
      TQrLabel(FF.FindComponent('RevLibre')).Top:=TopRev ;
      If CritEdt.BalBud.TotIniRev then
         BEGIN
         TQrLabel(FF.FindComponent('IRGen')).Top:=TopQIni ;
         TQrLabel(FF.FindComponent('QIniGen')).Top:=TopQRev ;
         TQrLabel(FF.FindComponent('QRevGen')).Top:=TopPlus ;
         TQrLabel(FF.FindComponent('IniGen')).Font.Style:=TQrLabel(FF.FindComponent('IniGen')).Font.Style+[fsItalic] ;
         TQrLabel(FF.FindComponent('RevGen')).Font.Style:=TQrLabel(FF.FindComponent('RevGen')).Font.Style+[fsItalic] ;
         END Else
         BEGIN
         TQrLabel(FF.FindComponent('QIniGen')).Top:=TopIni ;
         TQrLabel(FF.FindComponent('QRevGen')).Top:=TopQRev ;
         END ;
      TQrLabel(FF.FindComponent('IRLibre')).Top:=TQrLabel(FF.FindComponent('IRGen')).Top ;
      TQrLabel(FF.FindComponent('QIniLibre')).Top:=TQrLabel(FF.FindComponent('QIniGen')).Top ;
      TQrLabel(FF.FindComponent('QRevLibre')).Top:=TQrLabel(FF.FindComponent('QRevGen')).Top ;
      TQrLabel(FF.FindComponent('IniLibre')).Font.Style:=TQrLabel(FF.FindComponent('IniGen')).Font.Style ;
      TQrLabel(FF.FindComponent('RevLibre')).Font.Style:=TQrLabel(FF.FindComponent('RevGen')).Font.Style ;

      if CritEdt.BalBud.Qte then Largeur:=TQrLabel(FF.FindComponent('QRevGen')).Top+TQrLabel(FF.FindComponent('QRevGen')).Height+1 Else
         if CritEdt.BalBud.TotIniRev then Largeur:=TQrLabel(FF.FindComponent('IRGen')).Top+TQrLabel(FF.FindComponent('IRGen')).Height+1
                                     Else Largeur:=TQrLabel(FF.FindComponent('RevGen')).Top+TQrLabel(FF.FindComponent('RevGen')).Height+1 ;
      END ;
   BD.Height:=Largeur ;
   BDLibre.Height:=Largeur ;
  (* *)
   END ;
END ;


Type TNumLL = Array[1..5] Of Integer ;
Procedure QuelNumL(TL : TQRLabel ; ii : Integer ; Var ll : Integer ; Var NumLL : TNumLL) ;
BEGIN
TL.SynLigne:=0 ;
If TL.Visible Then
   BEGIN
   NumLL[ii]:=ll ; TL.SynLigne:=ll ; Inc(ll) ;
   END Else TL.SynLigne:=0 ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : ???
Créé le ...... : 01/01/1900
Modifié le ... : 18/10/2005
Description .. : 
Suite ........ : 
Suite ........ :   // FQ 16536 SBO 18/10/2005 Mise en place des options
Suite ........ : de révision dans les balances budgétaires
Mots clefs ... : 
*****************************************************************}
Procedure RetailleBandesBudgetes(FF : TForm ; CritEdt : TCritEdt ; LibRev : String ; BD, BDLibre, BDRecap, BDFPrim, BDFin : TQRBAND) ;
Const TopIni = 2 ; TopRev = 16 ; TopRea = 30 ; TopEca = 44 ;
Var I, Min, Max        : Byte ;
    Largeur            : Integer ;
    LaBudNat           : String ;
    C1, C2, C3, C4     : TComponent ;
    HIni, HRev, HRea, HEca : Integer ;
    VisIni, VisRev, VisRea, VisEca : Boolean ;
    Numll : TNumLL ;
    ll : Integer ;
    function _TestNat( vStCode : String ) : Boolean ;
      begin
      result := ( CritEdt.BalBud.NatureBud = '' )
               or ( CritEdt.BalBud.NatureBud = TraduireMemoire('<<Tous>>') )
               or ( pos( vStCode, CritEdt.BalBud.NatureBud ) > 0 ) ;
      end ;
BEGIN
Fillchar(Numll,SizeOf(Numll),#0) ; ll:=0 ;
Min:=CritEdt.BalBud.FormatPrint.ColMin ;
//Max:=CritEdt.BalBud.FormatPrint.ColMax-1 ;
Max:=CritEdt.BalBud.FormatPrint.ColMax-1 ;
LaBudNat := ''  ;
{ Bande Détail }
  TQrLabel(FF.FindComponent('IniGen')).Visible        := _TestNat('INI') or _TestNat('ANO');
  QuelNumL(TQrLabel(FF.FindComponent('IniGen')),1,ll,Numll) ;
  TQrLabel(FF.FindComponent('IniGenTotal')).Visible   := _TestNat('INI') or _TestNat('ANO');
  TQrLabel(FF.FindComponent('IniGenTotal')).SynLigne  := ll ;

  TQrLabel(FF.FindComponent('RevGen')).Visible        := _TestNat('DM')  ;
  QuelNumL(TQrLabel(FF.FindComponent('RevGen')),2,ll,Numll) ;
  TQrLabel(FF.FindComponent('RevGenTotal')).Visible   := _TestNat('DM')  ;
  TQrLabel(FF.FindComponent('RevGenTotal')).SynLigne  := ll ;

  TQrLabel(FF.FindComponent('ReaGen')).Visible        := CritEdt.BalBud.Realise ;
  QuelNumL(TQrLabel(FF.FindComponent('ReaGen')),3,ll,Numll) ;
  TQrLabel(FF.FindComponent('ReaGenTotal')).Visible   := CritEdt.BalBud.Realise ;
  TQrLabel(FF.FindComponent('ReaGenTotal')).SynLigne  := ll ;

  TQrLabel(FF.FindComponent('EcaGen')).Visible        := CritEdt.BalBud.Realise ;
  QuelNumL(TQrLabel(FF.FindComponent('EcaGen')),4,ll,Numll) ;
  TQrLabel(FF.FindComponent('EcaGenTotal')).Visible   := CritEdt.BalBud.Realise ;
  TQrLabel(FF.FindComponent('EcaGenTotal')).SynLigne  := ll ;

// Total Edition
  TQrLabel(FF.FindComponent('ReaTotal')).Visible   := CritEdt.BalBud.Realise ;
  TQrLabel(FF.FindComponent('TotTotalR')).Visible  := CritEdt.BalBud.Realise ;
  TQrLabel(FF.FindComponent('EcaTotal')).Visible   := CritEdt.BalBud.Realise ;
  TQrLabel(FF.FindComponent('TotTotalE')).Visible  := CritEdt.BalBud.Realise ;

// Total Pied Primaire en composite
if BDFPrim<>Nil then
   BEGIN
   TQrLabel(FF.FindComponent('ReaTotPrim')).Visible      := CritEdt.BalBud.Realise ;
   TQrLabel(FF.FindComponent('ReaTotPrimTotal')).Visible := CritEdt.BalBud.Realise ;
   TQrLabel(FF.FindComponent('EcaTotPrim')).Visible      := CritEdt.BalBud.Realise ;
   TQrLabel(FF.FindComponent('EcaTotPrimTotal')).Visible := CritEdt.BalBud.Realise ;
   END ;

IF not _TestNat('DM') then  // Nature du budget à INI UNIQUEMENT ! donc pas de révisé...
   BEGIN
   For i:=Min to Max do
       BEGIN
       C1:=FF.FindComponent('IniGenCum'+InttoStr(i))   ; TQRLabel(C1).Top:=TopIni ; TQRLabel(C1).Visible:=True ;
       C2:=FF.FindComponent('RevGenCum'+InttoStr(i))   ; TQRLabel(C2).Visible:=False ;
       C3:=FF.FindComponent('ReaGenCum'+InttoStr(i))   ; TQRLabel(C3).Visible:=CritEdt.BalBud.Realise ;
       C4:=FF.FindComponent('EcaGenCum'+InttoStr(i))   ; TQRLabel(C4).Visible:=CritEdt.BalBud.Realise ;
       if CritEdt.BalBud.Realise then
          BEGIN
          TQRLabel(C3).Top:=TopRev ;
          TQRLabel(C4).Top:=TopRea ;
          END Else
          BEGIN
          TQRLabel(C3).Top:=TopRea ;
          TQRLabel(C4).Top:=TopEca ;
          END ;
       END ;
   TQrLabel(FF.FindComponent('IniGenTotal')).Top:=TopIni ; TQrLabel(FF.FindComponent('IniGen')).Top:=TopIni ;
   if CritEdt.BalBud.Realise then
      BEGIN
      TQrLabel(FF.FindComponent('ReaGenTotal')).Top:=TopRev ;
      TQrLabel(FF.FindComponent('ReaGen')).Top:=TopRev ;
      TQrLabel(FF.FindComponent('EcaGenTotal')).Top:=TopRea ;
      TQrLabel(FF.FindComponent('EcaGen')).Top:=TopRea ;
      Largeur:=TQrLabel(FF.FindComponent('EcaGen')).Top+TQrLabel(FF.FindComponent('EcaGen')).Height+1 ;
      END Else
      BEGIN
      TQrLabel(FF.FindComponent('ReaGenTotal')).Top:=TopRea ;
      TQrLabel(FF.FindComponent('ReaGen')).Top:=TopRea ;
      TQrLabel(FF.FindComponent('EcaGenTotal')).Top:=TopEca ;
      TQrLabel(FF.FindComponent('EcaGen')).Top:=TopEca ;
      Largeur:=TQrLabel(FF.FindComponent('IniGen')).Top+TQrLabel(FF.FindComponent('IniGen')).Height+1 ;
      END ;
   BD.Height:=Largeur ;
   END Else
If not ( _TestNat('INI') or _TestNat('ANO') ) then // Nature du budget à DM X UNIQUEMENT ! donc pas de initial
   BEGIN
   For i:=Min to Max do
       BEGIN
       C1:=FF.FindComponent('IniGenCum'+InttoStr(i)); TQRLabel(C1).Visible:=False ;
       C2:=FF.FindComponent('RevGenCum'+InttoStr(i)); TQRLabel(C2).Top:=TopIni ; TQRLabel(C2).Visible:=True ;
       C3:=FF.FindComponent('ReaGenCum'+InttoStr(i)); TQRLabel(C3).Visible:=CritEdt.BalBud.Realise ;
       C4:=FF.FindComponent('EcaGenCum'+InttoStr(i)); TQRLabel(C4).Visible:=CritEdt.BalBud.Realise ;
       if CritEdt.BalBud.Realise then
          BEGIN
          TQRLabel(C3).Top:=TopRev ;
          TQRLabel(C4).Top:=TopRea ;
          END Else
          BEGIN
          TQRLabel(C3).Top:=TopRea ;
          TQRLabel(C4).Top:=TopEca ;
          END ;
       END ;
   TQrLabel(FF.FindComponent('RevGenTotal')).Top:=TopIni ;
   TQrLabel(FF.FindComponent('RevGen')).Top:=TopIni ;
   TQrLabel(FF.FindComponent('RevGen')).Caption:=LibRev+' '+LaBudNat ;
   if CritEdt.BalBud.Realise then
      BEGIN
      TQrLabel(FF.FindComponent('ReaGenTotal')).Top:=TopRev ;
      TQrLabel(FF.FindComponent('ReaGen')).Top:=TopRev ;
      TQrLabel(FF.FindComponent('EcaGenTotal')).Top:=TopRea ;
      TQrLabel(FF.FindComponent('EcaGen')).Top:=TopRea ;
      Largeur:=TQrLabel(FF.FindComponent('EcaGen')).Top+TQrLabel(FF.FindComponent('EcaGen')).Height+1 ;
      END Else
      BEGIN
      Largeur:=TQrLabel(FF.FindComponent('RevGen')).Top+TQrLabel(FF.FindComponent('RevGen')).Height+1 ;
      END ;
   BD.Height:=Largeur ;
   END Else
   BEGIN      { Nature du budget à TOUS }
   For i:=Min to Max do
       BEGIN
       C1:=FF.FindComponent('IniGenCum'+InttoStr(i)) ; TQRLabel(C1).Visible:=True ; TQRLabel(C1).Top:=TopIni ;
       C2:=FF.FindComponent('RevGenCum'+InttoStr(i)) ; TQRLabel(C2).Visible:=True ; TQRLabel(C2).Top:=TopRev ;
       C3:=FF.FindComponent('ReaGenCum'+InttoStr(i)) ; TQRLabel(C3).Visible:=CritEdt.BalBud.Realise ;
       C4:=FF.FindComponent('EcaGenCum'+InttoStr(i)) ; TQRLabel(C4).Visible:=CritEdt.BalBud.Realise ;
       TQRLabel(C3).Top:=TopRea ; TQRLabel(C4).Top:=TopEca ;
       END ;
   TQrLabel(FF.FindComponent('IniGenTotal')).Top:=TopIni ;
   TQrLabel(FF.FindComponent('IniGen')).Top:=TopIni ;
   TQrLabel(FF.FindComponent('RevGenTotal')).Top:=TopRev ;
   TQrLabel(FF.FindComponent('RevGen')).Top:=TopRev ;
   TQrLabel(FF.FindComponent('ReaGenTotal')).Top:=TopRea ;
   TQrLabel(FF.FindComponent('ReaGen')).Top:=TopRea ;
   TQrLabel(FF.FindComponent('EcaGenTotal')).Top:=TopEca ;
   TQrLabel(FF.FindComponent('EcaGen')).Top:=TopEca ;

   if CritEdt.BalBud.Realise then Largeur:=TopEca+TQrLabel(FF.FindComponent('EcaGen')).Height+1
                             Else Largeur:=TopRev+TQrLabel(FF.FindComponent('RevGen')).Height+1 ;
   BD.Height:=Largeur ;
   END ;
BDLibre.Height:=BD.Height ;
If BDRecap<>Nil then BDRecap.Height:=BD.Height ; {Recap en composite}
{Affichage ou pas du Total général en réalisé et ecart}
For i:=Min to Max do
    BEGIN
    C1:=FF.FindComponent('ReaTotal'+InttoStr(i)) ; TQRLabel(C1).Visible:=CritEdt.BalBud.Realise ;
    C2:=FF.FindComponent('EcaTotal'+InttoStr(i)) ; TQRLabel(C2).Visible:=CritEdt.BalBud.Realise ;
    (* Plante ici *)
    if BDFPrim<>Nil then { Label en Pied de compte Primaire en composite }
       BEGIN
       C3:=FF.FindComponent('ReaTotPrim'+InttoStr(i)) ; TQRLabel(C3).Visible:=CritEdt.BalBud.Realise ;
       C4:=FF.FindComponent('EcaTotPrim'+InttoStr(i)) ; TQRLabel(C4).Visible:=CritEdt.BalBud.Realise ;
       END ;
    END ;
if CritEdt.BalBud.Realise then Largeur:=TQrLabel(FF.FindComponent('EcaTotal')).Top+TQrLabel(FF.FindComponent('EcaTotal')).Height+5
                          Else Largeur:=TQrLabel(FF.FindComponent('BudTotal')).Top+TQrLabel(FF.FindComponent('BudTotal')).Height+5 ;

BDFin.Height:=Largeur ;
if BDFPrim<>Nil then BDFPrim.Height:=Largeur ; { Pied de compte Primaire en composite }

(* Retaille autres bandes selon BDetail**)
if (CritEdt.BalBud.RuptOnly<>Sans) then
   BEGIN
   { Bande Libre ou Rupture }
   VisIni:=TQrLabel(FF.FindComponent('IniGen')).Visible ; HIni:=TQrLabel(FF.FindComponent('IniGen')).Top ;
   VisRev:=TQrLabel(FF.FindComponent('RevGen')).Visible ; HRev:=TQrLabel(FF.FindComponent('RevGen')).Top ;
   VisRea:=TQrLabel(FF.FindComponent('ReaGen')).Visible ; HRea:=TQrLabel(FF.FindComponent('ReaGen')).Top ;
   VisEca:=TQrLabel(FF.FindComponent('EcaGen')).Visible ; HEca:=TQrLabel(FF.FindComponent('EcaGen')).Top ;
      (**)
   TQrLabel(FF.FindComponent('IniLibre')).Visible:=VisIni ; TQrLabel(FF.FindComponent('IniLibreTotal')).Visible:=VisIni ;
   TQrLabel(FF.FindComponent('IniLibre')).Top:=HIni ; TQrLabel(FF.FindComponent('IniLibreTotal')).Top:=HIni ;
   TQrLabel(FF.FindComponent('RevLibre')).Visible:=VisRev ; TQrLabel(FF.FindComponent('RevLibreTotal')).Visible:=VisRev ;
   TQrLabel(FF.FindComponent('RevLibre')).Top:=HRev ; TQrLabel(FF.FindComponent('RevLibreTotal')).Top:=HRev ;
   TQrLabel(FF.FindComponent('ReaLibre')).Visible:=VisRea ; TQrLabel(FF.FindComponent('ReaLibreTotal')).Visible:=VisRea ;
   TQrLabel(FF.FindComponent('ReaLibre')).Top:=HRea ; TQrLabel(FF.FindComponent('ReaLibreTotal')).Top:=HRea ;
   TQrLabel(FF.FindComponent('EcaLibre')).Visible:=VisEca ; TQrLabel(FF.FindComponent('EcaLibreTotal')).Visible:=VisEca ;
   TQrLabel(FF.FindComponent('EcaLibre')).Top:=HEca ; TQrLabel(FF.FindComponent('EcaLibreTotal')).Top:=HEca ;
   TQrLabel(FF.FindComponent('RevLibre')).Caption:=TQrLabel(FF.FindComponent('RevGen')).Caption ;
   { Bande Recap }
   If CritEdt.BalBud.AvecCptSecond and (BDRecap<>Nil) then
      BEGIN
      TQrLabel(FF.FindComponent('IniRecap')).Visible:=VisIni ; TQrLabel(FF.FindComponent('IniRecapTotal')).Visible:=VisIni ;
      TQrLabel(FF.FindComponent('IniRecap')).Top:=HIni       ; TQrLabel(FF.FindComponent('IniRecapTotal')).Top:=HIni ;
      TQrLabel(FF.FindComponent('RevRecap')).Visible:=VisRev ; TQrLabel(FF.FindComponent('RevRecapTotal')).Visible:=VisRev ;
      TQrLabel(FF.FindComponent('RevRecap')).Top:=HRev       ; TQrLabel(FF.FindComponent('RevRecapTotal')).Top:=HRev ;
      TQrLabel(FF.FindComponent('ReaRecap')).Visible:=VisRea ; TQrLabel(FF.FindComponent('ReaRecapTotal')).Visible:=VisRea ;
      TQrLabel(FF.FindComponent('ReaRecap')).Top:=HRea       ; TQrLabel(FF.FindComponent('ReaRecapTotal')).Top:=HRea ;
      TQrLabel(FF.FindComponent('EcaRecap')).Visible:=VisEca ; TQrLabel(FF.FindComponent('EcaRecapTotal')).Visible:=VisEca ;
      TQrLabel(FF.FindComponent('EcaRecap')).Top:=HEca       ; TQrLabel(FF.FindComponent('EcaRecapTotal')).Top:=HEca ;
      TQrLabel(FF.FindComponent('RevRecap')).Caption:=TQrLabel(FF.FindComponent('RevGen')).Caption ;
      END ;
   For i:=Min to Max do
       BEGIN
       VisIni:=TQRLabel(FF.FindComponent('IniGenCum'+InttoStr(i))).Visible ;
       Hini:=TQRLabel(FF.FindComponent('IniGenCum'+InttoStr(i))).Top ;
       VisRev:=TQRLabel(FF.FindComponent('RevGenCum'+InttoStr(i))).Visible ;
       HRev:=TQRLabel(FF.FindComponent('RevGenCum'+InttoStr(i))).Top ;
       VisRea:=TQRLabel(FF.FindComponent('ReaGenCum'+InttoStr(i))).Visible ;
       HRea:=TQRLabel(FF.FindComponent('ReaGenCum'+InttoStr(i))).Top ;
       VisEca:=TQRLabel(FF.FindComponent('EcaGenCum'+InttoStr(i))).Visible ;
       HEca:=TQRLabel(FF.FindComponent('EcaGenCum'+InttoStr(i))).Top ;

       TQRLabel(FF.FindComponent('IniLibre'+InttoStr(i))).Visible:=VisIni ;
       TQRLabel(FF.FindComponent('IniLibre'+InttoStr(i))).Top:=HIni ;
       TQRLabel(FF.FindComponent('RevLibre'+InttoStr(i))).Visible:=VisRev ;
       TQRLabel(FF.FindComponent('RevLibre'+InttoStr(i))).Top:=HRev ;
       TQRLabel(FF.FindComponent('ReaLibre'+InttoStr(i))).Visible:=VisRea;
       TQRLabel(FF.FindComponent('ReaLibre'+InttoStr(i))).Top:=HRea ;
       TQRLabel(FF.FindComponent('EcaLibre'+InttoStr(i))).Visible:=VisEca ;
       TQRLabel(FF.FindComponent('EcaLibre'+InttoStr(i))).Top:=HEca ;
       If CritEdt.BalBud.AvecCptSecond and (BDRecap<>Nil) then
          BEGIN
          TQRLabel(FF.FindComponent('IniRecap'+InttoStr(i))).Visible:=VisIni ;
          TQRLabel(FF.FindComponent('IniRecap'+InttoStr(i))).Top:=HIni ;
          TQRLabel(FF.FindComponent('RevRecap'+InttoStr(i))).Visible:=VisRev ;
          TQRLabel(FF.FindComponent('RevRecap'+InttoStr(i))).Top:=HRev ;
          TQRLabel(FF.FindComponent('ReaRecap'+InttoStr(i))).Visible:=VisRea;
          TQRLabel(FF.FindComponent('ReaRecap'+InttoStr(i))).Top:=HRea ;
          TQRLabel(FF.FindComponent('EcaRecap'+InttoStr(i))).Visible:=VisEca ;
          TQRLabel(FF.FindComponent('EcaRecap'+InttoStr(i))).Top:=HEca ;
          END ;
       END ;
   END ;
 (* Fin Selon**)
END ;

Procedure RetailleBandesBudEcarts(FF : TForm ; CritEdt : TCritEdt ; LibRev : String ; BD, BDLibre, BDRecap, BDFin : TQRBAND) ;
Const TopIni = 2 ; TopRev = 16 ; TopQIni = 30 ; TopQRev = 44 ; TopPlus = 58 ;
Var I, Min, Max        : Byte ;
    Largeur            : Integer ;
    LaBudNat           : String ;
    C1, C2, C3, C4, C5 : TComponent ;
    HIni, HRev, HQIni, HQRev, HIR : Integer ;
    VisIni, VisRev, VisQIni, VisQRev, VisIR : Boolean ;
    QueTotBud : Boolean ;
    Numll : TNumLL ;
    ll : Integer ;
BEGIN
Fillchar(Numll,SizeOf(Numll),#0) ; ll:=0 ;
Min:=CritEdt.BalBud.FormatPrint.ColMin ;
//Max:=CritEdt.BalBud.FormatPrint.ColMax-1 ;
Max:=CritEdt.BalBud.FormatPrint.ColMax ;
LaBudNat:=CritEdt.BalBud.NatureBud ;
QueTotBud:=CritEdt.BalBud.QueTotalBud ;
TQrLabel(FF.FindComponent('RevGen')).Caption:=LibRev ;

TQrLabel(FF.FindComponent('IniGen')).Visible:=((LaBudNat='') or (LaBudNat='INI')) And (Not QueTotBud) ;
QuelNumL(TQrLabel(FF.FindComponent('IniGen')),1,ll,Numll) ;
TQrLabel(FF.FindComponent('RevGen')).Visible:=((LaBudNat='') or (Copy(LaBudNat,1,2)='DM')) And (Not QueTotBud) ;
QuelNumL(TQrLabel(FF.FindComponent('RevGen')),2,ll,Numll) ;
TQrLabel(FF.FindComponent('QIniGen')).Visible:=((LaBudNat='') or (LaBudNat='INI')) and CritEdt.BalBud.Qte And (Not QueTotBud) ;
QuelNumL(TQrLabel(FF.FindComponent('QIniGen')),3,ll,Numll) ;
TQrLabel(FF.FindComponent('QRevGen')).Visible:=((LaBudNat='') or (Copy(LaBudNat,1,2)='DM')) and CritEdt.BalBud.Qte And (Not QueTotBud) ;
QuelNumL(TQrLabel(FF.FindComponent('QRevGen')),4,ll,Numll) ;
TQrLabel(FF.FindComponent('IRGen')).Visible:=CritEdt.BalBud.TotIniRev Or QueTotBud ;
QuelNumL(TQrLabel(FF.FindComponent('IRGen')),5,ll,Numll) ;

TQrLabel(FF.FindComponent('IniGen')).Font.Style:=TQrLabel(FF.FindComponent('IniGen')).Font.Style-[fsItalic] ;
TQrLabel(FF.FindComponent('RevGen')).Font.Style:=TQrLabel(FF.FindComponent('RevGen')).Font.Style-[fsItalic] ;

IF LaBudNat='INI' then  { Nature du budget à INI }
   BEGIN
   For i:=Min to Max do
       BEGIN
       C1:=FF.FindComponent('IniGenCum'+InttoStr(i))  ; TQRLabel(C1).Top:=TopIni ; TQRLabel(C1).Visible:=True ;
       C2:=FF.FindComponent('RevGenCum'+InttoStr(i))  ; TQRLabel(C2).Visible:=False ;
       C3:=FF.FindComponent('QIniGenCum'+InttoStr(i)) ; TQRLabel(C3).Visible:=CritEdt.BalBud.Qte ;
       TQRLabel(C1).Font.Style:=TQRLabel(C1).Font.Style-[fsItalic] ;
       if CritEdt.BalBud.Qte then TQRLabel(C3).Top:=TopRev
                             Else TQRLabel(C3).Top:=TopQIni ;
       C4:=FF.FindComponent('QRevGenCum'+InttoStr(i))   ; TQRLabel(C4).Visible:=False ;
       TQRLabel(C4).Top:=TopQRev ;
       END ;
   TQrLabel(FF.FindComponent('IniGen')).Top:=TopIni ;
   if CritEdt.BalBud.Qte then
      BEGIN
      TQrLabel(FF.FindComponent('QIniGen')).Top:=TopRev ;
      Largeur:=TQrLabel(FF.FindComponent('QIniGen')).Top+TQrLabel(FF.FindComponent('QIniGen')).Height+1 ;
      END Else
      BEGIN
      TQrLabel(FF.FindComponent('QIniGen')).Top:=TopQIni ;
      Largeur:=TQrLabel(FF.FindComponent('IniGen')).Top+TQrLabel(FF.FindComponent('IniGen')).Height+1 ;
      END ;
   END Else
If Copy(LaBudNat,1,2)='DM' then { Nature du budget à DM X }
   BEGIN
   For i:=1 to 10 do
       BEGIN
       C1:=FF.FindComponent('IniGenCum'+InttoStr(i))   ; TQRLabel(C1).Visible:=False ;
       C2:=FF.FindComponent('RevGenCum'+InttoStr(i))   ; TQRLabel(C2).Top:=TopIni ; TQRLabel(C2).Visible:=True ;
       C3:=FF.FindComponent('QIniGenCum'+InttoStr(i))   ; TQRLabel(C3).Visible:=False ;
       TQRLabel(C2).Font.Style:=TQRLabel(C2).Font.Style-[fsItalic] ;
       C4:=FF.FindComponent('QRevGenCum'+InttoStr(i))   ; TQRLabel(C4).Visible:=CritEdt.BalBud.Qte ;
       if CritEdt.BalBud.Qte then TQRLabel(C4).Top:=TopRev
                             Else TQRLabel(C4).Top:=TopQRev ;
       END ;
   TQrLabel(FF.FindComponent('RevGen')).Caption:=LibRev+' '+LaBudNat ;
   TQrLabel(FF.FindComponent('RevGen')).Top:=TopIni ;
   if CritEdt.BalBud.Qte then
      BEGIN
      TQrLabel(FF.FindComponent('QRevGen')).Top:=TopRev ;
      Largeur:=TQrLabel(FF.FindComponent('QRevGen')).Top+TQrLabel(FF.FindComponent('QRevGen')).Height+1 ;
      END Else
      BEGIN
      TQrLabel(FF.FindComponent('QRevGen')).Top:=TopQRev ;
      Largeur:=TQrLabel(FF.FindComponent('RevGen')).Top+TQrLabel(FF.FindComponent('RevGen')).Height+1 ;
      END ;
   END Else
   BEGIN      { Nature du budget à TOUS }
   For i:=1 to 10 do
       BEGIN
       C1:=FF.FindComponent('IniGenCum'+InttoStr(i))  ; TQRLabel(C1).Visible:=Not QueTotBud ; TQRLabel(C1).Top:=TopIni ;
       C2:=FF.FindComponent('RevGenCum'+InttoStr(i))  ; TQRLabel(C2).Visible:=Not QueTotBud  ; TQRLabel(C2).Top:=TopRev ;
       C3:=FF.FindComponent('QIniGenCum'+InttoStr(i)) ; TQRLabel(C3).Visible:=CritEdt.BalBud.Qte And (Not QueTotBud) ;
       C4:=FF.FindComponent('QRevGenCum'+InttoStr(i)) ; TQRLabel(C4).Visible:=CritEdt.BalBud.Qte And (Not QueTotBud) ;
       C5:=FF.FindComponent('IRGenCum'+InttoStr(i))   ; TQRLabel(C5).Visible:=CritEdt.BalBud.TotIniRev Or QueTotBud ;
       TQRLabel(C1).SynLigne:=NumLL[1] ;
       TQRLabel(C2).SynLigne:=NumLL[2] ;
       TQRLabel(C3).SynLigne:=NumLL[3] ;
       TQRLabel(C4).SynLigne:=NumLL[4] ;
       TQRLabel(C5).SynLigne:=NumLL[5] ;
       If QueTotBud Then { Positionnement du top des variables montants }
          BEGIN
          TQRLabel(C5).Top:=TopIni ;
          TQRLabel(C3).Top:=TopQRev ;
          TQRLabel(C4).Top:=TopPlus ;
          TQRLabel(C1).Font.Style:=TQRLabel(C1).Font.Style+[fsItalic] ;
          TQRLabel(C2).Font.Style:=TQRLabel(C2).Font.Style+[fsItalic] ;
          END Else
          If CritEdt.BalBud.TotIniRev then
             BEGIN
             TQRLabel(C5).Top:=TopQIni ;
             TQRLabel(C3).Top:=TopQRev ;
             TQRLabel(C4).Top:=TopPlus ;
             TQRLabel(C1).Font.Style:=TQRLabel(C1).Font.Style+[fsItalic] ;
             TQRLabel(C2).Font.Style:=TQRLabel(C2).Font.Style+[fsItalic] ;
             END Else
             BEGIN
             TQRLabel(C3).Top:=TopQIni ;
             TQRLabel(C4).Top:=TopQRev ;
             TQRLabel(C1).Font.Style:=TQRLabel(C1).Font.Style-[fsItalic] ;
             TQRLabel(C2).Font.Style:=TQRLabel(C2).Font.Style-[fsItalic] ;
             END ;
       END ;
   TQrLabel(FF.FindComponent('IniGen')).Top:=TopIni ;
   TQrLabel(FF.FindComponent('RevGen')).Top:=TopRev ;
   If QueTotBud Then
      BEGIN
      TQrLabel(FF.FindComponent('IRGen')).Top:=TopIni ;
      END Else
      If CritEdt.BalBud.TotIniRev then
         BEGIN
         TQrLabel(FF.FindComponent('IRGen')).Top:=TopQIni ;
         TQrLabel(FF.FindComponent('QIniGen')).Top:=TopQRev ;
         TQrLabel(FF.FindComponent('QRevGen')).Top:=TopPlus ;
         TQrLabel(FF.FindComponent('IniGen')).Font.Style:=TQrLabel(FF.FindComponent('IniGen')).Font.Style+[fsItalic] ;
         TQrLabel(FF.FindComponent('RevGen')).Font.Style:=TQrLabel(FF.FindComponent('RevGen')).Font.Style+[fsItalic] ;
         END Else
         BEGIN
         TQrLabel(FF.FindComponent('QIniGen')).Top:=TopIni ;
         TQrLabel(FF.FindComponent('QRevGen')).Top:=TopQRev ;
         END ;
   if CritEdt.BalBud.Qte then Largeur:=TQrLabel(FF.FindComponent('QRevGen')).Top+TQrLabel(FF.FindComponent('QRevGen')).Height+1 Else
      if CritEdt.BalBud.TotIniRev Or QueTotBud then Largeur:=TQrLabel(FF.FindComponent('IRGen')).Top+TQrLabel(FF.FindComponent('IRGen')).Height+1
                                               Else Largeur:=TQrLabel(FF.FindComponent('RevGen')).Top+TQrLabel(FF.FindComponent('RevGen')).Height+1 ;
   END ;
BD.Height:=Largeur ;
BDLibre.Height:=BD.Height ;
If BDRecap<>Nil then BDRecap.Height:=BD.Height ;

{Selon BDetail ...}
{Selon BDLibre ou Rupture ...}
if (CritEdt.BalBud.RuptOnly<>Sans) then
   BEGIN
   { Bande Libre ou Rupture }
   VisIni:=TQrLabel(FF.FindComponent('IniGen')).Visible ; HIni:=TQrLabel(FF.FindComponent('IniGen')).Top ;
   VisRev:=TQrLabel(FF.FindComponent('RevGen')).Visible ; HRev:=TQrLabel(FF.FindComponent('RevGen')).Top ;
   VisQini:=TQrLabel(FF.FindComponent('QIniGen')).Visible ; HQini:=TQrLabel(FF.FindComponent('QIniGen')).Top ;
   VisQRev:=TQrLabel(FF.FindComponent('QRevGen')).Visible ; HQrev:=TQrLabel(FF.FindComponent('QRevGen')).Top ;
   VisIR:=TQrLabel(FF.FindComponent('IRGen')).Visible ; HIR:=TQrLabel(FF.FindComponent('IRGen')).Top ;
      (**)
   TQrLabel(FF.FindComponent('IniLibre')).Visible:=VisIni ; TQrLabel(FF.FindComponent('IniLibre')).Top:=HIni ;
   TQrLabel(FF.FindComponent('IniLibre')).SynLigne:=NumLL[1] ;
   TQrLabel(FF.FindComponent('RevLibre')).Visible:=VisRev ; TQrLabel(FF.FindComponent('RevLibre')).Top:=HRev ;
   TQrLabel(FF.FindComponent('RevLibre')).SynLigne:=NumLL[2] ;
   TQrLabel(FF.FindComponent('QIniLibre')).Visible:=VisQini ; TQrLabel(FF.FindComponent('QiniLibre')).Top:=HQini ;
   TQrLabel(FF.FindComponent('QIniLibre')).SynLigne:=NumLL[3] ;
   TQrLabel(FF.FindComponent('QRevLibre')).Visible:=VisQrev ; TQrLabel(FF.FindComponent('QRevLibre')).Top:=HQrev ;
   TQrLabel(FF.FindComponent('QRevLibre')).SynLigne:=NumLL[4] ;
   TQrLabel(FF.FindComponent('IRLibre')).Visible:=VisIR ; TQrLabel(FF.FindComponent('IRLibre')).Top:=HIR ;
   TQrLabel(FF.FindComponent('IRLibre')).SynLigne:=NumLL[5] ;
   TQrLabel(FF.FindComponent('RevLibre')).Caption:=TQrLabel(FF.FindComponent('RevGen')).Caption ;
   { Bande Recap }
   If CritEdt.BalBud.AvecCptSecond and (BDRecap<>Nil) then
      BEGIN
      TQrLabel(FF.FindComponent('IniRecap')).Visible:=VisIni ; TQrLabel(FF.FindComponent('IniRecap')).Top:=HIni ;
      TQrLabel(FF.FindComponent('IniRecap')).SynLigne:=NumLL[1] ;
      TQrLabel(FF.FindComponent('RevRecap')).Visible:=VisRev ; TQrLabel(FF.FindComponent('RevRecap')).Top:=HRev ;
      TQrLabel(FF.FindComponent('RevRecap')).SynLigne:=NumLL[2] ;
      TQrLabel(FF.FindComponent('QIniRecap')).Visible:=VisQini ; TQrLabel(FF.FindComponent('QiniRecap')).Top:=HQini ;
      TQrLabel(FF.FindComponent('QIniRecap')).SynLigne:=NumLL[3] ;
      TQrLabel(FF.FindComponent('QRevRecap')).Visible:=VisQrev ; TQrLabel(FF.FindComponent('QRevRecap')).Top:=HQrev ;
      TQrLabel(FF.FindComponent('QRevRecap')).SynLigne:=NumLL[4] ;
      TQrLabel(FF.FindComponent('IRRecap')).Visible:=VisIR ; TQrLabel(FF.FindComponent('IRRecap')).Top:=HIR ;
      TQrLabel(FF.FindComponent('IRRecap')).SynLigne:=NumLL[5] ;
      TQrLabel(FF.FindComponent('RevRecap')).Caption:=TQrLabel(FF.FindComponent('RevGen')).Caption ;
      END ;
   For i:=Min to Max do
       BEGIN
       VisIni:=TQRLabel(FF.FindComponent('IniGenCum'+InttoStr(i))).Visible ;
       Hini:=TQRLabel(FF.FindComponent('IniGenCum'+InttoStr(i))).Top ;
       VisRev:=TQRLabel(FF.FindComponent('RevGenCum'+InttoStr(i))).Visible ;
       HRev:=TQRLabel(FF.FindComponent('RevGenCum'+InttoStr(i))).Top ;
       VisQini:=TQRLabel(FF.FindComponent('QIniGenCum'+InttoStr(i))).Visible ;
       HQini:=TQRLabel(FF.FindComponent('QIniGenCum'+InttoStr(i))).Top ;
       VisQRev:=TQRLabel(FF.FindComponent('QRevGenCum'+InttoStr(i))).Visible ;
       HQRev:=TQRLabel(FF.FindComponent('QRevGenCum'+InttoStr(i))).Top ;
       VisIR:=TQRLabel(FF.FindComponent('IRGenCum'+InttoStr(i))).Visible ;
       HIR:=TQRLabel(FF.FindComponent('IRGenCum'+InttoStr(i))).Top ;

       TQRLabel(FF.FindComponent('IniLibre'+InttoStr(i))).Visible:=VisIni ;
       TQRLabel(FF.FindComponent('IniLibre'+InttoStr(i))).Top:=HIni ;
       TQRLabel(FF.FindComponent('IniLibre'+InttoStr(i))).SynLigne:=Numll[1] ;
       TQRLabel(FF.FindComponent('RevLibre'+InttoStr(i))).Visible:=VisRev ;
       TQRLabel(FF.FindComponent('RevLibre'+InttoStr(i))).Top:=HRev ;
       TQRLabel(FF.FindComponent('RevLibre'+InttoStr(i))).SynLigne:=Numll[2] ;
       TQRLabel(FF.FindComponent('QIniLibre'+InttoStr(i))).Visible:=VisQini;
       TQRLabel(FF.FindComponent('QIniLibre'+InttoStr(i))).Top:=HQini ;
       TQRLabel(FF.FindComponent('QIniLibre'+InttoStr(i))).SynLigne:=Numll[3] ;
       TQRLabel(FF.FindComponent('QRevLibre'+InttoStr(i))).Visible:=VisQRev;
       TQRLabel(FF.FindComponent('QRevLibre'+InttoStr(i))).Top:=HQRev ;
       TQRLabel(FF.FindComponent('QRevLibre'+InttoStr(i))).SynLigne:=Numll[4] ;
       TQRLabel(FF.FindComponent('IRLibre'+InttoStr(i))).Visible:=VisIR ;
       TQRLabel(FF.FindComponent('IRLibre'+InttoStr(i))).Top:=HIR ;
       TQRLabel(FF.FindComponent('IRLibre'+InttoStr(i))).SynLigne:=Numll[5] ;
       If CritEdt.BalBud.AvecCptSecond and (BDRecap<>Nil) then
          BEGIN
          TQRLabel(FF.FindComponent('IniRecap'+InttoStr(i))).Visible:=VisIni ;
          TQRLabel(FF.FindComponent('IniRecap'+InttoStr(i))).Top:=HIni ;
          TQRLabel(FF.FindComponent('IniRecap'+InttoStr(i))).SynLigne:=Numll[1] ;
          TQRLabel(FF.FindComponent('RevRecap'+InttoStr(i))).Visible:=VisRev ;
          TQRLabel(FF.FindComponent('RevRecap'+InttoStr(i))).Top:=HRev ;
          TQRLabel(FF.FindComponent('RevRecap'+InttoStr(i))).SynLigne:=Numll[2] ;
          TQRLabel(FF.FindComponent('QIniRecap'+InttoStr(i))).Visible:=VisQini;
          TQRLabel(FF.FindComponent('QIniRecap'+InttoStr(i))).Top:=HQIni ;
          TQRLabel(FF.FindComponent('QIniRecap'+InttoStr(i))).SynLigne:=Numll[3] ;
          TQRLabel(FF.FindComponent('QRevRecap'+InttoStr(i))).Visible:=VisQRev;
          TQRLabel(FF.FindComponent('QRevRecap'+InttoStr(i))).Top:=HQRev ;
          TQRLabel(FF.FindComponent('QRevRecap'+InttoStr(i))).SynLigne:=Numll[4] ;
          TQRLabel(FF.FindComponent('IRRecap'+InttoStr(i))).Visible:=VisIR ;
          TQRLabel(FF.FindComponent('IRRecap'+InttoStr(i))).Top:=HIR ;
          TQRLabel(FF.FindComponent('IRRecap'+InttoStr(i))).SynLigne:=Numll[5] ;
          END ;
       END ;
   END ;

(* *)
END ;


end.
