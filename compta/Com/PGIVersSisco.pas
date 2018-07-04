unit PGIVersSisco;

interface
uses
  SysUtils, Classes, Hctrls, StdCtrls, Ent1,
  hmsgbox,HStatus,HEnt1, Controls,
  {$IFNDEF EAGLCLIENT}
{$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  Sisco, Recordcomsis, UtilTrans,
  CritEdt, Recordcom, UTOB;

Procedure ExportEcriture (Where : string; Listb : TListBox ; var Fichier : TextFile; CritEdt : TCritEdt; sanscoll : Boolean; What,Stat : string; Qualifp:string='N'; VersQuadra : Boolean=FALSE);
Function TransfertPgiVersSisco(StFichier : String; Listb : TListBox; Trans : TFTransfertcom; sanscoll : Boolean;c,f : string; Qualifp:string='N'; VersQuadra : Boolean=FALSE; datearrete : string='') : Boolean ;
procedure Ecriture_Collectif(var Fichier : TextFile; OkEcritureFichier: Boolean = TRUE);
procedure Ecriture_par3(var Fichier : TextFile; Stat : string);
Procedure FaitSISCOCptTiers(Var FF : TextFile ; Var CritEdt : TCritEdt) ;
Procedure ExportBalance (Where : string; Listb : TListBox ; var Fichier : TextFile; CritEdt : TCritEdt; sanscoll : Boolean;Trans : TFTransfertcom; datearrete : string='');
procedure Ecriture_par4(var Fichier : TextFile;Trans : TFTransfertcom);
Procedure Ecriture_jrl(var Fichier : TextFile; Where : string) ;
procedure ExportAnouveau (TOBANO,TOBLET : TOB; sanscoll : Boolean;var Fichier : TextFile; CritEdt : TCritEdt);
procedure Ecriture_Stat(var Fichier : TextFile; Section : string);
Function TransferversQuadra(StFichier : String; Listb : TListBox; Trans : TFTransfertcom; sanscoll : Boolean; c,f : string; Qualifp:string='N'; VersQuadra : Boolean=FALSE) : Boolean ;

implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  {$ENDIF MODENT1}
  Paramsoc;


var
AncienPeriode, AncJournal : string;
Carfour, Cercli           : string;
TOBTiers                  : TOB;
AjoutCar, OkQCar          : Boolean;
const
           Formatjournal = 'J%2.2s%20.20s%10.10s%5.5s%-3.3s';

Function TransferversQuadra(StFichier : String; Listb : TListBox; Trans : TFTransfertcom; sanscoll : Boolean; c,f : string; Qualifp:string='N'; VersQuadra : Boolean=FALSE) : Boolean ;
Var
Q1   : TQuery;
Ind  : integer;
St,FsQ : string;
TQ     : TOB;
I      : integer;
begin
    try
     Ind := 1;    Result := FALSE; AjoutCar := TRUE; OkQCar := FALSE;
     Q1 := OpenSql ('SELECT * from EXERCICE ORDER BY EX_EXERCICE',FALSE);
     TQ := TOB.Create ('', nil, -1);
     TQ.LoadDetailDB('EXERCICE','','',Q1,False,True) ;
     Ferme (Q1);
     For i := 0 to TQ.detail.count -1 do
     begin
        Trans.Dateecr1         := StrToDate(TQ.detail[i].GetValue ('EX_DATEDEBUT'));
        Trans.Dateecr2         := StrToDate(TQ.detail[i].GetValue ('EX_DATEFIN'));
        Trans.Exo              := TQ.detail[i].GetValue ('EX_EXERCICE');
        St := StFichier;
        FsQ := ReadTokenPipe (St, '.') + '_' + IntToStr(ind)+'.TRT';
        TransfertPgiVersSisco(FsQ, Listb, Trans, sanscoll, c, f, Qualifp, VersQuadra);
        inc (ind);
     end;
    TQ.free;
    Finally
           Result := TRUE;
    end;
end;

Function TransfertPgiVersSisco(StFichier : String; Listb : TListBox; Trans : TFTransfertcom; sanscoll : Boolean; c,f : string; Qualifp:string='N'; VersQuadra : Boolean=FALSE; datearrete : string='') : Boolean ;
Var
       Fichier          : TextFile ;
       CritEdt          : TCritEdt;
       Q1               : TQuery;
       Where,Wherej     : string;
       Jourx            : string;
       savlg            : integer;
       Ccli, Cfou       : integer;
BEGIN
     Fillchar(CritEdt,SizeOf(CritEdt),#0) ;
     CritEdt.Complement := Trans.SuppCarAux;
     Carfour := f; Cercli := c;
     Savlg := VH^.Cpta[fbGene].Lg;
     Ccli := 0; Cfou := 0;
     Q1 := Opensql ('select count(t_auxiliaire) from tiers where t_natureauxi="CLI" and '+
        '(t_auxiliaire not like "' +Cercli+ '%")', TRUE);
     if not Q1.EOF then
        Ccli := Q1.Fields[0].asinteger;
     ferme (Q1);
     Q1 := Opensql ('select count(t_auxiliaire) from tiers where t_natureauxi="FOU" and '+
        '(t_auxiliaire not like "' +Carfour+ '%")', TRUE);
     if not Q1.EOF then
        Cfou := Q1.Fields[0].asinteger;
     ferme (Q1);
{$IFNDEF EAGLSERVER}
     if VersQuadra and ((Ccli > 0) or (Cfou > 0)) then
     begin
       if  not OkQCar then
       begin
          if not (PGIAsk ('Confirmez-vous  l''ajout des caractères '+Cercli + ' ou '+ Carfour +' sur tous les comptes Tiers', 'Paramétrage Sisco II ')=mrYes) then
          begin
           Ccli :=0; Cfou := 0;
           AjoutCar := FALSE;
          end
          else AjoutCar := TRUE;
          OkQCar := TRUE;
       end
       else
       begin
           if not AjoutCar then begin Ccli :=0; Cfou := 0; end;
       end;
     end;
{$ENDIF}

     if (savlg < 10) and (not sanscoll)  and ((Ccli <> 0) or (Cfou <> 0))then
        VH^.Cpta[fbGene].Lg := VH^.Cpta[fbGene].Lg+1;

     if VH^.Cpta[fbGene].Lg > 10 then
     begin
      PgiInfo ('Longueur > 10, transfert impossible', 'Transfert ComSx');
      Result := FALSE; exit ;
     end;

     AncienPeriode := ''; AncJournal := '';
     if Trans.Exo <> '' then
             Where := 'E_EXERCICE="'+Trans.Exo+'"';
     if Trans.Jr <> ''  then
     begin
               Jourx := Trans.Jr;
               if Where <> '' then  Where := Where + ' AND ';
               Where := Where + '('+ RendCommandeJournal('E',Jourx)+') ';
     end;
     if Trans.Etabi <> '' then
     begin
               if Where <> '' then  Where := Where + ' AND ';
              // Where := Where + '(E_ETABLISSEMENT="'+ Trans.Etabi+'") ';
               Where := Where + RendCommandeComboMulti('E',Trans.Etabi,'ETABLISSEMENT');
     end;
     if (Trans.Dateecr1 <> iDate1900) and (Trans.Dateecr2 <> iDate1900) then
     begin
               if Where <> '' then  Where := Where + ' AND';
               Where := Where + ' (E_DATECOMPTABLE>="'+USDateTime(Trans.Dateecr1)+'"' +
                              ' AND E_DATECOMPTABLE<="'+USDateTime(Trans.Dateecr2)+'")'
     end
     else
     begin
              if (Trans.Dateecr1 <> iDate1900) then
              begin
                   if Where <> '' then  Where := Where + ' AND';
                   Where := Where + ' (E_DATECOMPTABLE>="'+USDateTime(Trans.Dateecr1)+'")';
              end else
              if (Trans.Dateecr2 <> iDate1900) then
              begin
                   if Where <> '' then  Where := Where + ' AND';
                   Where := Where + ' (E_DATECOMPTABLE<="'+USDateTime(Trans.Dateecr2)+'")';
              end;

     end ;

     (*
         CritEdt.Exo.Deb := Trans.Dateecr1;
         CritEdt.Exo.Fin := Trans.Dateecr2;
     *)
     Q1 := OpenSql ('SELECT EX_DATEDEBUT,EX_DATEFIN FROM EXERCICE WHERE EX_EXERCICE="'+
     Trans.Exo +'"', FALSE);
     if not Q1.EOF then
     begin
          CritEdt.Exo.Deb := Q1.FindField ('EX_DATEDEBUT').asDatetime;
          CritEdt.Exo.Fin := Q1.FindField ('EX_DATEFIN').asDateTime;
     end;
     Ferme (Q1);
     CritEdt.Decimale := V_PGI.OkDecV;
     If VH^.TenueEuro Then
        CritEdt.Monnaie := 2;

     AssignFile(Fichier,StFichier) ; Rewrite(Fichier) ;
     AlimCptResultat(CritEdt) ;
     FaitSISCODebut(Fichier) ;

     // ajout me 9-12-2002 SISCO EXO
     if Trans.Typearchive = 'EXE' then
        FaitSISCO00(Fichier, CritEdt, 'E')
     else
     if Trans.Typearchive = 'JRL' then
        FaitSISCO00(Fichier, CritEdt, 'J')
     else
     if Trans.Typearchive = 'BAL' then
        FaitSISCO00(Fichier, CritEdt, 'B');

     Ecriture_par3(Fichier, Trans.psection);
     Ecriture_par4(Fichier, Trans);
     if Trans.Typearchive <> 'BAL' then
     begin
          Wherej := '';
          if Trans.Jr <> ''  then
          begin
               Jourx := Trans.Jr;
               Wherej := ' Where ' + RendCommandeJournal('J',Jourx);
          end;
          Ecriture_jrl(Fichier, Wherej) ;
     end;
     if not sanscoll then
        Ecriture_Collectif(Fichier);

     if (Ccli = 0) and (Cfou = 0) then
     begin
          Carfour := ''; Cercli := '';
     end;

     if (Trans.psection = 'TL') or (Trans.psection = 'ANA') then
     Ecriture_Stat(Fichier, Trans.psection);

     FaitSISCOCptGen(Fichier, CritEdt, False) ;
     if not sanscoll then
        FaitSISCOCptTiers(Fichier, CritEdt) ;
//FaitSISCOPER(Fichier,CritEdt.DateFin,CritEdt.Exo) ;
     if Trans.balance then
     begin
//        CritEdt.Exo := Trans.Dateecr2;
        ExportBalance (Where, Listb, Fichier, CritEdt, sanscoll, Trans, datearrete);
     end
     else
         ExportEcriture (Where, Listb, Fichier, CritEdt, sanscoll, Trans.Typearchive,Trans.psection, Qualifp, VersQuadra);

     FaitSISCOFin(Fichier) ;
     FiniMove ;
     CloseFile(Fichier) ;
     VH^.Cpta[fbGene].Lg := savlg;

    Result := TRUE;
END;

Function OkDoublon (Where,JANO : string) : Boolean;
var
Q  : Tquery;
begin
            Result := TRUE;
{$IFDEF PB}
            Q:=OpenSQL('Select count(*),E_DEVISE,E_ETABLISSEMENT,E_GENERAL,E_AUXILIAIRE  from ecriture  Where '+ Where + ' and (E_JOURNAL="'+JANO+'")'+
               ' AND (E_QUALIFPIECE="N" or E_QUALIFPIECE="") and E_ETATLETTRAGE<>"RI" and E_ETATLETTRAGE<>"" '+
               ' Group by E_DEVISE,E_ETABLISSEMENT,E_GENERAL,E_AUXILIAIRE ', TRUE);
              While not Q.EOF do
              begin
               if Q.Fields[0].asinteger > 1 then
               begin
                PGIInfo ('Il existe des écritures ayant plusieurs devises, sur le journal '+GetParamSocSecur('SO_JALOUVRE', 'ANO')+' Transfert ComSx');
                Result := FALSE;
                ferme(Q);
                exit;
               end;
               Q.next;
              end;
              ferme(Q);
{$ENDIF}
            Q:=OpenSQL('Select  DISTINCT E_DEVISE from ecriture  Where '+ Where + ' and (E_JOURNAL="'+JANO+'")'+
               ' AND (E_QUALIFPIECE="N" or E_QUALIFPIECE="") and E_ETATLETTRAGE<>"RI" and E_ETATLETTRAGE<>"" '+
               ' Group by E_DEVISE,E_ETABLISSEMENT,E_GENERAL,E_AUXILIAIRE ', TRUE);
            if Q.RecordCount > 1 then
            begin
                PGIInfo ('Il existe des écritures ayant plusieurs devises, sur le journal '+GetParamSocSecur('SO_JALOUVRE', 'ANO')+' Transfert ComSx');
                Result := FALSE;
                ferme(Q);
                exit;
            end;
            ferme(Q);
            Q:=OpenSQL('Select  DISTINCT E_ETABLISSEMENT from ecriture  Where '+ Where + ' and (E_JOURNAL="'+JANO+'")'+
               ' AND (E_QUALIFPIECE="N" or E_QUALIFPIECE="") and E_ETATLETTRAGE<>"RI" and E_ETATLETTRAGE<>"" '+
               ' Group by E_DEVISE,E_ETABLISSEMENT,E_GENERAL,E_AUXILIAIRE ', TRUE);
            if Q.RecordCount > 1 then
            begin
                PGIInfo ('Il existe des écritures ayant plusieurs établissement, sur le journal '+GetParamSocSecur('SO_JALOUVRE', 'ANO')+' Transfert ComSx');
                Result := FALSE;
                ferme(Q);
                exit;
            end;
            ferme(Q);
end;


Procedure ExportEcriture (Where : string; Listb : TListBox ; var Fichier : TextFile; CritEdt : TCritEdt; sanscoll : Boolean; What,Stat : string; Qualifp:string='N'; VersQuadra : Boolean=FALSE);
var
Q,Q1,Q2                  : TQuery;
Parc_ecr                 : ^ar_ecr;
Orderby                  : string;
Wheres                   : string;
NbLigne                  : integer;
NumFolio                 : integer;
Dtt                      : TDatetime;
YY,MM,JJ                 : Word ;
MD,MC                    : double;
Ligne                    : string;
natureaux                : string;
T1                       : TOB;
cpte,JANO                : string;
TOBJ,TOBANO              : TOB;
Libelle,cpteauto         : string;
nature                   : string;
Oklettre                 : Boolean;
TOBLettrage              : TOB;
Whereana                 : string;
begin
    TOBANO := nil;
    TOBJ := TOB.Create('', nil, -1);
    Q:=OpenSQL('SELECT J_JOURNAL,J_LIBELLE,J_CONTREPARTIE,J_NATUREJAL FROM JOURNAL',TRUE) ;

    TOBJ.LoadDetailDB('JOURNAL', '', '', Q, TRUE, FALSE);
    Ferme(Q);

//    CritEdt.Exo.fin := VH^.EnCours.Fin;
    CritEdt.Decimale := V_PGI.OkDecV;
    Orderby := ' order by E_EXERCICE, E_PERIODE, E_JOURNAL, E_NUMEROPIECE, E_QUALIFPIECE,E_NUMLIGNE, E_NUMECHE, E_AUXILIAIRE';
    if Where <> '' then
         Wheres := 'Where '+ Where + ' AND (E_QUALIFPIECE="'+Qualifp+'" or E_QUALIFPIECE="")';

    if What = 'EXE' then // cas exercice
    begin
         JANO := GetParamSocSecur('SO_JALOUVRE', 'ANO');
         Q:=OpenSQL('SELECT * FROM ECRITURE '+ Wheres +
         ' and E_JOURNAL="'+JANO+'" '+ OrderBy,TRUE);
         if not Q.EOF then
         begin
               if Q.FindField('E_ECRANOUVEAU').asstring = 'H' then
               begin
                    Ferme (Q);
                    Oklettre  := OkLettrage (Where, Listb);
                     if (not Oklettre) and (not VersQuadra)then
                     begin
                       PGIBox ('Le lettrage n''est pas équilibré sur cet exercice, aucune écriture ne sera exportée','Export');
                       if not sanscoll then TOBTiers.free;
                       Ferme (Q);
                       exit;
                     end;
               end else
               begin
                    if not OkDoublon (Where,JANO) then
                    begin
                       if not sanscoll then TOBTiers.free;
                       Ferme (Q);
                       exit;
                    end;

                    TOBANO := TOB.Create('', nil, -1);
                    TOBLettrage := TOB.Create('', nil, -1);
                    TOBANO.LoadDetailDB('ECRITURE', '', '', Q, TRUE, FALSE);
                    Ferme(Q);
                    FaitSISCOPER(Fichier, CritEdt.Exo.deb, CritEdt.Exo) ;
                    Ligne := Format(Formatjournal,['AA', 'Journal aa', '', '','AA']);
                    WriteLn(Fichier, Ligne) ;
                    //WriteLn(Fichier,'JAAJournal aa') ;
                    FaitSISCOFolio(Fichier,1) ;
                    ExportAnouveau (TOBANO,TOBLettrage, sanscoll, Fichier, CritEdt);
                    Oklettre  := OkLettrage (Where, Listb, TOBLettrage);
                    if (not Oklettre) and (not VersQuadra) then
                    begin
                       PGIBox ('Le lettrage n''est pas équilibré sur cet exercice, aucune écriture ne sera exportée','Export');
                       if not sanscoll then TOBTiers.free;
                       TOBLettrage.free;
                       exit;
                    end;
                    TOBLettrage.free;

               end;
         end else
         begin
              Ferme (Q);
              if GetParamSocSecur ('SO_EXOV8', '') <> '' then
              begin
                    Ferme (Q);
                    Oklettre  := OkLettrage (Where, Listb);
                    if not Oklettre then
                    begin
                       PGIBox ('Attention, le lettrage n''est pas équilibré sur cet exercice, aucune écriture ne sera exportée','Export');
                       if not sanscoll then TOBTiers.free;
                       exit;
                    end;
              end;
         end;
    end;

    Q1:=OpenSQL('SELECT * FROM ECRITURE '+ Wheres + OrderBy,TRUE);
    if Q1.EOF then
    begin
         PGIInfo ('Il n''y a pas d''écritures pour les critères de selection choisies','');
         Ferme(Q1) ;
         if not sanscoll then TOBTiers.free;
         exit;
    end;
    NbLigne:=0 ; NumFolio:=1 ;

    While Not Q1.Eof Do
    BEGIN
      Dtt := Q1.FindField ('E_DATECOMPTABLE').asDateTime;
     CritEdt.Exo.fin := Dtt;
     if What = 'EXE' then // cas exercice
     begin
          if (Q1.FindField ('E_ECRANOUVEAU').asstring ='OAN') then
          begin Q1.next; continue; end;
     end;
     if (AncienPeriode <> Q1.FindField ('E_PERIODE').asstring) then
     begin
           FaitSISCOPER(Fichier, CritEdt.Exo.fin, CritEdt.Exo) ;
           AfficheListeCom( FormatDateTime (Traduitdateformat('mmmm yyyy'),Q1.FindField ('E_DATECOMPTABLE').asdatetime),Listb);
           if AncJournal = Q1.FindField ('E_JOURNAL').asstring then
           begin
              T1 := TOBJ.FindFirst(['J_JOURNAL'], [AncJournal], FALSE);
              if T1 <> nil then
              begin
                   Libelle := T1.GetValue('J_LIBELLE');
                   cpteauto := T1.GetValue('J_CONTREPARTIE');
              end;
              WriteLn(Fichier,'J'+AncJournal+Libelle+cpteauto) ;
              FaitSISCOFolio(Fichier,NumFolio) ;
           end;
      end;
      AncienPeriode := Q1.FindField ('E_PERIODE').asstring;

      if AncJournal <> Q1.FindField ('E_JOURNAL').asstring then
      begin
          AncJournal := Q1.FindField ('E_JOURNAL').asstring;
          T1 := TOBJ.FindFirst(['J_JOURNAL'], [AncJournal], FALSE);
          if T1 <> nil then
          begin
               Libelle := T1.GetValue('J_LIBELLE');
               cpteauto := T1.GetValue('J_CONTREPARTIE');
          end;
           Ligne := Format(Formatjournal,[AncJournal, Libelle, cpteauto, '',AncJournal]);
           WriteLn(Fichier, Ligne) ;

//           WriteLn(Fichier,'J'+AncJournal+Libelle+cpteauto) ;
           FaitSISCOFolio(Fichier,NumFolio) ;
           AfficheListeCom('Ecriture du journal : '+ Q1.FindField ('E_JOURNAL').asstring, Listb);
      end;
      AncJournal := Q1.FindField ('E_JOURNAL').asstring;

      If NbLigne>=90 Then BEGIN NbLigne:=1 ; Inc(NumFolio) ; FaitSISCOFolio(Fichier,NumFolio) ; END ;
      DecodeDate(Dtt,YY,MM,JJ) ;
      New (Parc_ecr);
      Parc_ecr^.jour := FormatFloat('00', JJ);
      if (Q1.FindField ('E_AUXILIAIRE').asstring <> '')  and (not sanscoll) then
      begin
          T1 := TOBTiers.FindFirst(['T_AUXILIAIRE'], [Q1.FindField ('E_AUXILIAIRE').asstring], FALSE);
          if T1 <> nil then begin nature := T1.Getvalue ('T_NATUREAUXI'); end;

          if T1 <> nil then
          begin
               if (nature = 'CLI') or (nature = 'FOU') then
               begin
                   if T1.GetValue('T_NATUREAUXI') = 'CLI' then natureaux := CerCli
                   else
                   if T1.GetValue('T_NATUREAUXI') = 'FOU' then natureaux := Carfour ;
                   cpte := AGauche(Q1.FindField ('E_AUXILIAIRE').asstring,10,'0');
                   ControleCompte(Cpte);
                   if natureaux <> '' then
                      Cpte := natureaux+Cpte;
                   if CritEdt.Complement then
                   begin
                      cpte := BourreLess(Cpte, fbAux);
                      Cpte := AGauche(Cpte, 10, ' ');
                   end
                   else
                      Cpte := AGauche(Cpte, 10, '0');
                   if natureaux <> '' then
                      Parc_ecr^.Cpte := natureaux + Copy (cpte, 2, length(Cpte)-1)
                   else
                      Parc_ecr^.Cpte := Cpte;
              end
              else
              Parc_ecr^.Cpte := AGauche(Q1.FindField ('E_GENERAL').asstring,10,'0');
          end
          else
               Parc_ecr^.Cpte := AGauche(Q1.FindField ('E_AUXILIAIRE').asstring,10,'0');
      end
      else
         Parc_ecr^.Cpte := AGauche(Q1.FindField ('E_GENERAL').asstring,10,'0');
      Parc_ecr^.lib := Q1.FindField ('E_LIBELLE').asstring;
      Parc_ecr^.piece :=  Q1.FindField ('E_REFINTERNE').asstring;

     // ajout me 9-12-2002 SISCO EXO
       if (Q1.FindField ('E_ETATLETTRAGE').asstring = 'PL') or
       (Q1.FindField ('E_LETTRAGE').asstring = '')  or  (What <> 'EXE') then
          Parc_ecr^.lettre := ''
       else
       begin
          if Copy (Q1.FindField ('E_LETTRAGE').asstring,4,1) <> '' then
             Parc_ecr^.lettre := Copy (Q1.FindField ('E_LETTRAGE').asstring,4,1)
          else
             Parc_ecr^.lettre := Copy (Q1.FindField ('E_LETTRAGE').asstring,0,1);
          Parc_ecr^.suitelettre := Copy (Q1.FindField ('E_LETTRAGE').asstring,3,1);
       end;


      MD := Q1.FindField ('E_DEBIT').asfloat;
      MC := Q1.FindField ('E_CREDIT').asfloat;
      if MC > MD then
      begin
           MC := MC-MD; MD := 0.0;
      end
      else
      begin
           if MD > MC then
           begin
                MD := MD-MC; MC := 0.0;
           end;
      end;

      if ((arrondi (MD, V_PGI.OKDECV)) > 0)  then
      begin
           Parc_ecr^.debit := ADroite(StrfMontantSISCO(MD,13,CritEdt.Decimale,'',False),13,'0');
           Parc_ecr^.credit := ADroite('0',13,'0');
      end
      else
      begin
           Parc_ecr^.debit := ADroite('0',13,'0');
           Parc_ecr^.credit := ADroite(StrfMontantSISCO(MC,13,CritEdt.Decimale,'',False),13,'0');
      end;

      Parc_ecr^.echeance := FormatDateTime(Traduitdateformat('ddmmyy'),Q1.FindField ('E_DATEECHEANCE').AsDateTime);
      if Q1.FindField ('E_DATEECHEANCE').AsDateTime = iDate1900 then Parc_ecr^.echeance := '';
      Parc_ecr^.mode := Q1.FindField ('E_MODEPAIE').asstring;
      Parc_ecr^.quan := ADroite(StrfMontantSISCO(Q1.FindField ('E_QTE1').asfloat,13,CritEdt.Decimale,'',False), 8, '0');
      Parc_ecr^.Nocheque := Q1.FindField ('E_NUMTRAITECHQ').asstring;
      Parc_ecr^.date_cre := FormatDateTime(Traduitdateformat('ddmmyy'),Q1.FindField ('E_DATECOMPTABLE').AsDateTime);
      if Q1.FindField ('E_DEVISE').asstring = 'FRF' then
      begin
         Parc_ecr^.m_euro := 'F';
      end
      else
      begin
          if Q1.FindField ('E_DEVISE').asstring = 'EUR' then
             Parc_ecr^.m_euro := 'E'
          else
             Parc_ecr^.m_euro := V_PGI.DevisePivot;
      end;
      Parc_ecr^.mteuro := ADroite('0', 13, '0');
     // ajout me 9-12-2002 SISCO EXO
      if Parc_ecr^.lettre <> '' then
         Parc_ecr^.identecr := FormatDateTime(Traduitdateformat('ddmmyyyy'), Q1.FindField ('E_DATEPAQUETMIN').asdatetime);

     // ajout me 15-07-2003
     if Stat = 'TL' then Parc_ecr^.stat :=  Q1.FindField ('E_TABLE0').asstring
     else
     if Stat = 'ANA' then
     begin
       Whereana := ' Where Y_JOURNAL="'+Q1.FindField ('E_JOURNAL').asstring+'" AND Y_DATECOMPTABLE="'+
        USDateTime (Q1.findField('E_DATECOMPTABLE').AsDatetime) +
        '" AND Y_NUMEROPIECE=' + IntToStr(Q1.FindField ('E_NUMEROPIECE').asinteger) +
        ' AND Y_NUMLIGNE='+ IntToStr(Q1.findField('E_NUMLIGNE').Asinteger);
       Q2:=OpenSQL('SELECT Y_SECTION FROM ANALYTIQ '+ Whereana, TRUE);
       if not Q2.EOF then
        Parc_ecr^.stat := Copy( Q2.FindField ('Y_SECTION').asstring, 1, 4);
       Ferme(Q2);
     end;

      Ligne := Format(Formatmvtsis, [ Parc_ecr^.jour, Parc_ecr^.cpte,
        Parc_ecr^.lib, Parc_ecr^.piece, Parc_ecr^.lettre, Parc_ecr^.extra,
        Parc_ecr^.debit, Parc_ecr^.credit, Parc_ecr^.stat,
        Parc_ecr^.echeance, Parc_ecr^.mode, Parc_ecr^.quan, Parc_ecr^.date_cre,
        Parc_ecr^.topee, Parc_ecr^.Nocheque,
        Parc_ecr^.date_rel, Parc_ecr^.niv_rel, Parc_ecr^.devise, Parc_ecr^.mtdev,
        Parc_ecr^.date_valeur, Parc_ecr^.suitelettre, Parc_ecr^.m_euro,
        Parc_ecr^.mteuro, Parc_ecr^.identecr, Parc_ecr^.autre_monn]);
      writeln(Fichier, Ligne);
      if Parc_ecr <> nil then Dispose (Parc_ecr);
      Q1.Next;
    END;

    Ferme (Q1);
if not sanscoll then TOBTiers.free;
if TOBANO <> nil then  TOBANO.free;
if TOBJ <> nil then TOBj.free;
end;

procedure EcritureAnouveau(TOBLec : TOB; sanscoll : Boolean; var Fichier : TextFile);
var
T1                       : TOB;
Parc_ecr                 : ^ar_ecr;
Dtt                      : TDateTime;
YY,MM,JJ                 : Word ;
nature,cpte              : string;
natureaux,Ligne          : string;
MD,MC                    : double;
begin
                   Dtt := TOBLec.getvalue('E_DATECOMPTABLE');
                   DecodeDate(Dtt,YY,MM,JJ) ;
                   New (Parc_ecr);
                   Parc_ecr^.jour := FormatFloat('00', JJ);
                   if (TOBLec.getvalue ('E_AUXILIAIRE') <> '')  and (not sanscoll) then
                   begin
                        T1 := TOBTiers.FindFirst(['T_AUXILIAIRE'], [TOBLec.getvalue ('E_AUXILIAIRE')], FALSE);
                        if T1 <> nil then begin nature := T1.Getvalue ('T_NATUREAUXI'); end;

                        if T1 <> nil then
                        begin
                             if (nature = 'CLI') or (nature = 'FOU') then
                             begin
                                  if T1.GetValue('T_NATUREAUXI') = 'CLI' then natureaux := CerCli
                                  else
                                  if T1.GetValue('T_NATUREAUXI') = 'FOU' then natureaux := Carfour ;
                                  cpte := AGauche(TOBLec.getvalue('E_AUXILIAIRE'),10,'0');
                                  ControleCompte(Cpte);
                                  Cpte := natureaux+Cpte;
                                  Cpte := AGauche(Cpte,10,'0');
                                  if natureaux <> '' then
                                     Parc_ecr^.Cpte := natureaux + Copy (cpte, 2, length(Cpte)-1)
                                  else
                                     Parc_ecr^.Cpte := Cpte;
                             end
                             else
                             Parc_ecr^.Cpte := AGauche(TOBLec.getvalue('E_GENERAL'),10,'0');
                        end
                        else
                             Parc_ecr^.Cpte := AGauche(TOBLec.getvalue('E_AUXILIAIRE'),10,'0');
                   end
                   else
                   Parc_ecr^.Cpte := AGauche(TOBLec.getvalue ('E_GENERAL'),10,'0');
                   Parc_ecr^.lib := TOBLec.getvalue ('E_LIBELLE');
                   Parc_ecr^.piece :=  TOBLec.getvalue ('E_REFINTERNE');
                   if (TOBLec.getvalue ('E_ETATLETTRAGE') = 'PL') or
                   (TOBLec.getvalue ('E_LETTRAGE') = '') then
                      Parc_ecr^.lettre := ''
                   else
                   begin
                      if Copy (TOBLec.getvalue ('E_LETTRAGE'),4,1) <> '' then
                         Parc_ecr^.lettre := Copy (TOBLec.getvalue ('E_LETTRAGE'),4,1)
                      else
                         Parc_ecr^.lettre := Copy (TOBLec.getvalue ('E_LETTRAGE'),0,1);
                      Parc_ecr^.suitelettre := Copy (TOBLec.getvalue ('E_LETTRAGE'),3,1);
                   end;
                   MD := TOBLec.getvalue ('E_DEBIT');
                   MC := TOBLec.getvalue ('E_CREDIT');
                   if MC > MD then
                   begin
                         MC := MC-MD; MD := 0.0;
                   end
                   else
                   begin
                         if MD > MC then
                         begin
                              MD := MD-MC; MC := 0.0;
                         end;
                   end;

                   if ((arrondi (MD, V_PGI.OKDECV)) > 0)  then
                   begin
                        Parc_ecr^.debit := ADroite(StrfMontantSISCO(MD,13,V_PGI.OkDecV,'',False),13,'0');
                        Parc_ecr^.credit := ADroite('0',13,'0');
                   end
                   else
                   begin
                        Parc_ecr^.debit := ADroite('0',13,'0');
                        Parc_ecr^.credit := ADroite(StrfMontantSISCO(MC,13,V_PGI.OkDecV,'',False),13,'0');
                   end;

                   Parc_ecr^.echeance := FormatDateTime(Traduitdateformat('ddmmyy'),TOBLec.getvalue ('E_DATEECHEANCE'));
                   if TOBLec.getvalue ('E_DATEECHEANCE') = iDate1900 then Parc_ecr^.echeance := '';
                   Parc_ecr^.mode := TOBLec.getvalue ('E_MODEPAIE');
                   Parc_ecr^.quan := ADroite(StrfMontantSISCO(TOBLec.getvalue('E_QTE1'),13,V_PGI.OkDecV,'',False), 8, '0');
                   Parc_ecr^.Nocheque :=TOBLec.getvalue('E_NUMTRAITECHQ');
                   Parc_ecr^.date_cre := FormatDateTime(Traduitdateformat('ddmmyy'),TOBLec.getvalue ('E_DATECOMPTABLE'));
                   if TOBLec.getvalue('E_DEVISE') = 'FRF' then
                        Parc_ecr^.m_euro := 'F'
                   else
                   begin
                        if TOBLec.getvalue ('E_DEVISE') = 'EUR' then
                           Parc_ecr^.m_euro := 'E'
                        else
                           Parc_ecr^.m_euro := V_PGI.DevisePivot;
                   end;
                   Parc_ecr^.mteuro := ADroite('0', 13, '0');
                   Parc_ecr^.identecr := FormatDateTime(Traduitdateformat('ddmmyyyy'),TOBLec.getvalue ('E_DATEPAQUETMIN'));

                   Ligne := Format(Formatmvtsis, [ Parc_ecr^.jour, Parc_ecr^.cpte,
                         Parc_ecr^.lib, Parc_ecr^.piece, Parc_ecr^.lettre, Parc_ecr^.extra,
                         Parc_ecr^.debit, Parc_ecr^.credit, Parc_ecr^.stat,
                         Parc_ecr^.echeance, Parc_ecr^.mode, Parc_ecr^.quan, Parc_ecr^.date_cre,
                         Parc_ecr^.topee, Parc_ecr^.Nocheque,
                         Parc_ecr^.date_rel, Parc_ecr^.niv_rel, Parc_ecr^.devise, Parc_ecr^.mtdev,
                         Parc_ecr^.date_valeur, Parc_ecr^.suitelettre, Parc_ecr^.m_euro,
                         Parc_ecr^.mteuro, Parc_ecr^.identecr, Parc_ecr^.autre_monn]);
                   writeln(Fichier, Ligne);
                   if Parc_ecr <> nil then Dispose (Parc_ecr);
end;

procedure AJoutLesMontantSolde (TOBLET : TOB; CritEdt : TCritEdt; Etab, lStDevise : string;  var Fichier : TextFile; sanscoll : Boolean );
var
lStSQL : String;
Q      : TQuery;
TN1, TL: TOB;
id     : integer;
Compte : string;
vTobCptSolde : TOB;
begin
// équilibre de lettrage
// exemple de requête SELECT * FROM ECRITURE WHERE E_EXERCICE >= "004" AND E_DATECOMPTABLE < "20060101 00:00:00" AND E_ETABLISSEMENT = "001" AND E_DEVISE = "EUR" AND ((E_ETATLETTRAGE = "TL" AND E_DATEPAQUETMAX >= "20060101 00:00:00")) AND (E_GENERAL||E_AUXILIAIRE NOT IN (SELECT E_GENERAL||E_AUXILIAIRE FROM ECRITURE WHERE E_JOURNAL = "ANO" AND E_DATECOMPTABLE = "20060101 00:00:00"))
    TN1 := Tob.Create ('', nil , -1);
    lStSql := 'SELECT * FROM ECRITURE WHERE ' +
              'E_EXERCICE >= "' + GetParamSocSecur('SO_EXOV8', '001') + '" AND ' +
              'E_DATECOMPTABLE < "' + USDateTime(CritEdt.Exo.Deb) + '" AND ' +
              'E_ETABLISSEMENT = "' + Etab + '" AND ' +
              'E_DEVISE = "' + lStDevise + '" AND ' +
              //'(E_ETATLETTRAGE = "AL" OR ' +
              '((E_ETATLETTRAGE = "TL" AND E_DATEPAQUETMAX >= "' + USDateTime(CritEdt.Exo.Deb) + '")) AND ' +
              '(E_GENERAL||E_AUXILIAIRE NOT IN ' +
              '(SELECT E_GENERAL||E_AUXILIAIRE FROM ECRITURE WHERE E_JOURNAL = "'+ GetParamSocSecur('SO_JALOUVRE', 'ANO') + '" AND ' +
              'E_DATECOMPTABLE = "' + USDateTime(CritEdt.Exo.Deb) +'")) ' ;
    Q:=OpenSQL(lStSQl, TRUE);
    TN1.LoadDetailDB('ECRITURE', '', '', Q, TRUE, FALSE);
    Ferme (Q);
   for id := 0 to TN1.detail.Count - 1 do
   begin
       EcritureAnouveau(TN1.detail[id], sanscoll, Fichier);
       if (TN1.detail[id].GetValue('E_LETTRAGE') <> '') and
       (TN1.detail[id].GetValue('E_ETATLETTRAGE') <> 'PL') and
       (TN1.detail[id].GetValue('E_ETATLETTRAGE') <> 'RI') then
       begin
             if TOBLET = nil then TOBLET := TOB.Create('', nil, -1);
             if (TN1.detail[id].getvalue ('E_AUXILIAIRE') <> '') then
                Compte := TN1.detail[id].getvalue('E_AUXILIAIRE')
             else
                Compte:= TN1.detail[id].getvalue ('E_GENERAL');
             if Compte <> '' then
             begin
                   TL := TOBLET.FindFirst(['COMPTE'], [Compte], FALSE);
                   if TL = nil then
                   begin
                        TL := TOB.Create ('',TOBLET,-1);
                        TL.AddChampSupValeur('COMPTE', Compte);
                        TL.AddChampSupValeur('DEBIT', TN1.detail[id].GetValue( 'E_DEBIT'));
                        TL.AddChampSupValeur('CREDIT',  TN1.detail[id].GetValue('E_CREDIT'));
                   end
                   else
                   begin
                        TL.putvalue('DEBIT',  TL.Getvalue('DEBIT')+TN1.detail[id].GetValue('E_DEBIT'));
                        TL.putvalue('CREDIT',  TL.Getvalue('CREDIT')+TN1.detail[id].GetValue('E_CREDIT'));
                   end;
             end;
        end;
   end;
   TN1.free;
// Ajustement d'anouveau
//SELECT E_GENERAL, E_AUXILIAIRE, E_DEVISE, E_ETABLISSEMENT, SUM(E_DEBIT)-SUM(E_CREDIT) TOTAL, SUM(E_DEBITDEV)-SUM(E_CREDITDEV) TOTALDEV FROM ECRITURE LEFT JOIN GENERAUX ON G_GENERAL = E_GENERAL WHERE ((G_COLLECTIF = "-" AND G_LETTRABLE = "X") OR (G_COLLECTIF = "X")) AND E_DATECOMPTABLE < "20060101 00:00:00" AND E_ETABLISSEMENT = "001" AND E_DEVISE = "EUR" AND (E_ETATLETTRAGE = "AL" OR (E_ETATLETTRAGE = "TL" AND E_DATEPAQUETMAX >= "20060101 00:00:00")) AND (E_GENERAL||E_AUXILIAIRE NOT IN (SELECT E_GENERAL||E_AUXILIAIRE FROM ECRITURE WHERE E_JOURNAL = "ANO" AND E_DATECOMPTABLE = "20060101 00:00:00")) GROUP BY E_GENERAL, E_AUXILIAIRE, E_DEVISE, E_ETABLISSEMENT
   lStSql := 'SELECT E_GENERAL, E_AUXILIAIRE, E_DEVISE, E_ETABLISSEMENT, ' +
              'SUM(E_DEBIT) TOTDEBIT, SUM(E_CREDIT) TOTCREDIT ' +
              'FROM ECRITURE LEFT JOIN GENERAUX ON G_GENERAL = E_GENERAL WHERE ' +
              '((G_COLLECTIF = "-" AND G_LETTRABLE = "X") OR (G_COLLECTIF = "X")) AND ' +
              'E_DATECOMPTABLE < "' + USDateTime(CritEdt.Exo.Deb) + '" AND ' +
              'E_ETABLISSEMENT = "' + Etab + '" AND ' +
              'E_DEVISE = "' + lStDevise + '" AND E_QUALIFPIECE = "N" AND ' +
              '((E_ETATLETTRAGE = "TL" AND E_DATEPAQUETMAX >= "' + USDateTime(CritEdt.Exo.Deb) + '")) AND ' +
              '(E_GENERAL||E_AUXILIAIRE NOT IN ' +
              '(SELECT E_GENERAL||E_AUXILIAIRE FROM ECRITURE WHERE E_JOURNAL = "'+ GetParamSocSecur('SO_JALOUVRE', 'ANO') + '" AND ' +
              'E_DATECOMPTABLE = "' + USDateTime(CritEdt.Exo.Deb)  +'")) ' +
              'GROUP BY E_GENERAL, E_AUXILIAIRE, E_DEVISE, E_ETABLISSEMENT';

    vTobCptSolde := Tob.Create('', nil, -1);
    Q:=OpenSQL(lStSQl, TRUE);
    vTobCptSolde.LoadDetailDB('ECRITURE', '', '', Q, TRUE, FALSE);

   for id := 0 to vTobCptSolde.detail.Count - 1 do
   begin
          TN1 :=  vTobCptSolde.Detail[id];
          TN1.AddchampSupValeur('E_ETATLETTRAGE', 'AL', True);
          TN1.AddchampSupValeur('G_LETTRABLE', 'X', True);
          TN1.AddChampSupValeur('E_NUMEROPIECE', IntToStr(1), True);
          if TN1.getvalue ('TOTCREDIT') > 0 then
            TN1.putvalue('E_DEBIT', TN1.getvalue('TOTCREDIT'));
          if TN1.getvalue ('TOTDEBIT') > 0 then
            TN1.putvalue('E_CREDIT', TN1.getvalue('TOTDEBIT'));
          EcritureAnouveau(TN1, sanscoll, Fichier);
    end;
    FreeAndNil(vTobCptSolde);

end;

procedure ExportAnouveau (TOBANO,TOBLET : TOB; sanscoll : Boolean;var Fichier : TextFile; CritEdt : TCritEdt);
var
ii,id                    : integer;
TOBLec,TL                : TOB;
Where                    : string;
TOBN1                    : TOB;
Compte                   : string;
Cpte, Cpte2              : string;
Q                        : TQuery;
TobN,TobN2               : TOB;
begin
         TobN := TOB.create('', nil, -1);
         TobN.Dupliquer(TOBANO, TRUE, True);

         for ii := 0 to TOBANO.detail.Count - 1 do
         begin
              TOBLec := TOBANO.detail[ii];
              if (TOBLec.getvalue ('E_ETATLETTRAGE') = 'RI') or
              (TOBLec.getvalue ('E_ETATLETTRAGE') = '') then
              begin
                   EcritureAnouveau(TOBANO.detail[ii], sanscoll, Fichier);
              end // etatlettrage
              else
              begin
                   TOBN1 := TOB.Create('', nil, -1);
                   Where :=
                   ' Where (E_JOURNAL<>"'+GetParamSocSecur('SO_JALFERME', 'ANO')+'") and (E_ECRANOUVEAU<>"OAN") '
                    + ' and (E_DATECOMPTABLE<"'+USDateTime(CritEdt.Exo.Deb)+'" '
                    + ' and E_DATECOMPTABLE>="'+USDateTime(VH^.ExoV8.Deb)+'" )'
                    + ' and E_DEVISE="'+ TOBLec.getvalue ('E_DEVISE')+ '"'
                    + ' and E_ETABLISSEMENT="'+ TOBLec.getvalue ('E_ETABLISSEMENT')+ '"';
                   if TOBLec.getvalue ('E_AUXILIAIRE') <> '' then
                   Where := Where + ' and E_AUXILIAIRE="'+ TOBLec.getvalue ('E_AUXILIAIRE')+'"';
                   if TOBLec.getvalue ('E_GENERAL') <> '' then
                   Where := Where + ' and E_GENERAL="'+ TOBLec.getvalue ('E_GENERAL')+'"';

                   Where := Where + ' and ((E_DATEPAQUETMAX>="'+USDateTime(CritEdt.Exo.Deb)+'"' +
                   ' and E_ETATLETTRAGE="TL")'+
                   ' or E_ETATLETTRAGE="AL" or E_ETATLETTRAGE="PL")';

                  if (TOBLec.getvalue ('E_AUXILIAIRE') <> '') and (cpte = TOBLec.getvalue ('E_AUXILIAIRE')) then
                   begin
                        cpte := TOBLec.getvalue ('E_AUXILIAIRE');
                        // cas deux auxiliaires identiques mais cpte général différent
                        TobN2 := TobN.FindFirst(['E_AUXILIAIRE', 'E_GENERAL'], [TOBLec.getvalue ('E_AUXILIAIRE'), Cpte2], FALSE);
                        if TobN2 <> nil then
                        begin
                             TOBN2.free;
                             continue;
                        end;
                   end
                   else
                        cpte := TOBLec.getvalue ('E_AUXILIAIRE');
                   Cpte2 := TOBLec.getvalue ('E_GENERAL');

                   Q:=OpenSQL('SELECT * FROM ECRITURE '+ Where,TRUE);
                   TOBN1.LoadDetailDB('ECRITURE', '', '', Q, TRUE, FALSE);
                   Ferme(Q);

                   for id := 0 to TOBN1.detail.Count - 1 do
                   begin
                       EcritureAnouveau(TOBN1.detail[id], sanscoll, Fichier);
                       if (TOBN1.detail[id].GetValue('E_LETTRAGE') <> '') and
                       (TOBN1.detail[id].GetValue('E_ETATLETTRAGE') <> 'PL') and
                       (TOBN1.detail[id].GetValue('E_ETATLETTRAGE') <> 'RI') then
                       begin
                             if (TOBN1.detail[id].getvalue ('E_AUXILIAIRE') <> '') then
                                Compte := TOBN1.detail[id].getvalue('E_AUXILIAIRE')
                             else
                                Compte:= TOBN1.detail[id].getvalue ('E_GENERAL');
                             if Compte <> '' then
                             begin
                                   TL := TOBLET.FindFirst(['COMPTE'], [Compte], FALSE);
                                   if TL = nil then
                                   begin
                                        TL := TOB.Create ('',TOBLET,-1);
                                        TL.AddChampSupValeur('COMPTE', Compte);
                                        TL.AddChampSupValeur('DEBIT', TOBN1.detail[id].GetValue( 'E_DEBIT'));
                                        TL.AddChampSupValeur('CREDIT',  TOBN1.detail[id].GetValue('E_CREDIT'));
                                   end
                                   else
                                   begin
                                        TL.putvalue('DEBIT',  TL.Getvalue('DEBIT')+TOBN1.detail[id].GetValue('E_DEBIT'));
                                        TL.putvalue('CREDIT',  TL.Getvalue('CREDIT')+TOBN1.detail[id].GetValue('E_CREDIT'));
                                   end;
                             end;
                        end;
                   end;
                   TOBN1.free;
              end;
              TobN2 := TobN.FindFirst(['E_AUXILIAIRE', 'E_GENERAL'], [TOBLec.getvalue ('E_AUXILIAIRE'), TOBLec.getvalue ('E_GENERAL')], FALSE);
              if TobN2 <> nil then TOBN2.free;

         end;
         TobN.free;
         AJoutLesMontantSolde (TOBLET, CritEdt, TOBANO.detail[0].getvalue ('E_ETABLISSEMENT'), TOBANO.detail[0].getvalue ('E_DEVISE'), Fichier, sanscoll );

end;

Procedure ExportBalance (Where : string; Listb : TListBox ; var Fichier : TextFile; CritEdt : TCritEdt; sanscoll : Boolean;Trans : TFTransfertcom; datearrete : string='');
var
Q1        : TQuery;
Parc_ecr  : ^ar_ecr;
SQl       : string;
NbLigne   : integer;
NumFolio  : integer;
Dtt       : TDatetime;
YY,MM,JJ  : Word ;
MD,MC     : double;
Ligne     : string;
natureaux : string;
T1        : TOB;
cpte      : string;
EAUXILIAIRE : string;
cpteaux     : string;
Etabliss    : string;
begin

    if sanscoll then
     EAUXILIAIRE := ''
    else
     EAUXILIAIRE := 'E_AUXILIAIRE,';
     cpteaux := '';

    CritEdt.Decimale := V_PGI.OkDecV;
    if Trans.Etabi <> '' then //Etabliss := 'AND (E_ETABLISSEMENT="'+ Trans.Etabi+'") '
       Etabliss := 'AND ('+ RendCommandeComboMulti('E',Trans.Etabi,'ETABLISSEMENT')+') '
    else Etabliss := '';

    Sql := 'SELECT E_GENERAL,'+ EAUXILIAIRE + 'E_EXERCICE, ' +
     'sum(E_DEBIT) as TOTDEBIT, sum(E_CREDIT) as TOTCREDIT, E_QUALIFPIECE,'+
     ' G_NATUREGENE,G_LIBELLE FROM ECRITURE LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL ' +
     'WHERE  E_DATECOMPTABLE>="'+USDateTime(Trans.Dateecr1)+'" AND E_DATECOMPTABLE<="'+USDateTime(Trans.Dateecr2) +
     '" AND E_EXERCICE="'+Trans.Exo+'" AND (E_QUALIFPIECE="N" ) ' +
     Etabliss +
     ' GROUP BY E_GENERAL,'+ EAUXILIAIRE+ 'E_EXERCICE,E_QUALIFPIECE,G_NATUREGENE,G_LIBELLE '+
     ' ORDER BY E_GENERAL,'+ EAUXILIAIRE+'E_EXERCICE,E_QUALIFPIECE,G_NATUREGENE,G_LIBELLE';

    Q1:=OpenSQL(Sql,TRUE);
    if Q1.EOF then
    begin
         PGIInfo ('Il n''y a pas d''écritures pour les critères de selection choisies','');
         Ferme(Q1) ;
         if not sanscoll then TOBTiers.free;
         exit;
    end;
    NbLigne:=0 ; NumFolio:=1 ;
    if datearrete <> '' then
       FaitSISCOPER(Fichier, StrToDate(datearrete), CritEdt.Exo)
    else
       FaitSISCOPER(Fichier, CritEdt.Exo.fin, CritEdt.Exo) ;
    AfficheListeCom( FormatDateTime (Traduitdateformat('mmmm yyyy'),Trans.Dateecr2),Listb);

    WriteLn(Fichier,'JCOJOURNAL CO') ;
    FaitSISCOFolio(Fichier,NumFolio) ;
    AfficheListeCom('Ecriture du journal : CO', Listb);

    While Not Q1.Eof Do
    BEGIN
      if datearrete <> '' then
       Dtt := StrToDate(datearrete)
      else
       Dtt := Trans.Dateecr2;

      CritEdt.Exo.fin := Dtt;
      If NbLigne>=90 Then BEGIN NbLigne:=1 ; Inc(NumFolio) ; FaitSISCOFolio(Fichier,NumFolio) ; END ;
      DecodeDate(Dtt,YY,MM,JJ) ;
      New (Parc_ecr);
      Parc_ecr^.jour := FormatFloat('00', JJ);
      if EAUXILIAIRE <> '' then cpteaux := Q1.FindField ('E_AUXILIAIRE').asstring;

      if (cpteaux <> '')  and (not sanscoll) then
      begin
          T1 := TOBTiers.FindFirst(['T_AUXILIAIRE'], [Q1.FindField ('E_AUXILIAIRE').asstring], FALSE);
          if T1 <> nil then
          begin
               if T1.GetValue('T_NATUREAUXI') = 'CLI' then natureaux := CerCli
               else
               if T1.GetValue('T_NATUREAUXI') = 'FOU' then natureaux := Carfour ;
               cpte := AGauche(Q1.FindField ('E_AUXILIAIRE').asstring,10,'0');
               ControleCompte(Cpte);
               if natureaux <> '' then
                  Cpte := natureaux+Cpte;
               if CritEdt.Complement then
               begin
                      cpte := BourreLess(Cpte, fbAux);
                      Cpte := AGauche(Cpte, 10, ' ');
               end
               else
                      Cpte := AGauche(Cpte, 10, '0');

               if natureaux <> '' then
                  Parc_ecr^.Cpte := natureaux + Copy (cpte, 2, length(Cpte)-1)
               else
                  Parc_ecr^.Cpte := Cpte;
          end
          else
               Parc_ecr^.Cpte := AGauche(Q1.FindField ('E_AUXILIAIRE').asstring,10,'0');
      end
      else
         Parc_ecr^.Cpte := AGauche(Q1.FindField ('E_GENERAL').asstring,10,'0');
      Parc_ecr^.lib := Q1.FindField ('G_LIBELLE').asstring;
//      Parc_ecr^.piece :=  AGauche(IntToStr(Q1.FindField ('E_NUMEROPIECE').asinteger),10,'0');
      Parc_ecr^.lettre := '';
      MD := 0.0; MC := 0.0;
      if Q1.FindField ('TOTDEBIT').asFloat > Q1.FindField ('TOTCREDIT').asFloat then
      MD := arrondi (Q1.FindField ('TOTDEBIT').asFloat - Q1.FindField ('TOTCREDIT').asFloat,V_PGI.OKDECV)
      else
      MC := arrondi (Q1.FindField ('TOTCREDIT').asFloat - Q1.FindField ('TOTDEBIT').asFloat,V_PGI.OKDECV);

      if (MD > 0) then
      begin
           Parc_ecr^.debit := ADroite(StrfMontantSISCO(MD,13,CritEdt.Decimale,'',False),13,'0');
           Parc_ecr^.credit := ADroite('0',13,'0');
      end
      else
      begin
           Parc_ecr^.debit := ADroite('0',13,'0');
           Parc_ecr^.credit := ADroite(StrfMontantSISCO(MC,13,CritEdt.Decimale,'',False),13,'0');
      end;
      Parc_ecr^.echeance := '';
      Parc_ecr^.mode := '';
      Parc_ecr^.quan := ADroite(StrfMontantSISCO(0,13,CritEdt.Decimale,'',False), 8, '0');
      Parc_ecr^.Nocheque := '';
      if datearrete = '' then  // fiche 10433
         Parc_ecr^.date_cre := FormatDateTime(Traduitdateformat('ddmmyy'),VH^.Encours.Fin)
      else
         Parc_ecr^.date_cre := FormatDateTime(Traduitdateformat('ddmmyy'), StrToDate(datearrete));

      Parc_ecr^.m_euro := V_PGI.DevisePivot;
      Parc_ecr^.mteuro := ADroite('0', 13, '0');
      Ligne := Format(Formatmvtsis, [ Parc_ecr^.jour, Parc_ecr^.cpte,
        Parc_ecr^.lib, Parc_ecr^.piece, Parc_ecr^.lettre, Parc_ecr^.extra,
        Parc_ecr^.debit, Parc_ecr^.credit, Parc_ecr^.stat,
        Parc_ecr^.echeance, Parc_ecr^.mode, Parc_ecr^.quan, Parc_ecr^.date_cre,
        Parc_ecr^.topee, Parc_ecr^.Nocheque,
        Parc_ecr^.date_rel, Parc_ecr^.niv_rel, Parc_ecr^.devise, Parc_ecr^.mtdev,
        Parc_ecr^.date_valeur, Parc_ecr^.suitelettre, Parc_ecr^.m_euro,
        Parc_ecr^.mteuro, Parc_ecr^.identecr, Parc_ecr^.autre_monn]);
      writeln(Fichier, Ligne);
      if Parc_ecr <> nil then Dispose (Parc_ecr);
      Q1.Next;
    END;

    Ferme (Q1);
if not sanscoll then TOBTiers.free;
end;

procedure Ecriture_Collectif(var Fichier : TextFile; OkEcritureFichier: Boolean = TRUE);
var
Q1        : TQuery;
Pcol      : ^ar_colreg;
Ligne     : string;
Tcor      : TOB;
Cpte      : string;
begin
  if not OkEcritureFichier then
  begin
       Carfour := '0';
       Cercli := '9';
  end;

  Q1 := OpenSql ('SELECT * from CORRESP Where CR_TYPE="SIS"', TRUE);
  if Q1.EOF then
  begin
            Tcor := TOB.Create('CORRESP', Nil, -1);
            New(Pcol);
            Cpte := GetParamSocSecur ('SO_DEFCOLFOU', '4010000000');
            if Copy (Cpte,0,2) <> '40' then Cpte := '4010000000';
            Pcol^.cpte := AGauche(Cpte,10,'0');
            Pcol^.inf := Carfour+'0000000';
            Pcol^.sup := Carfour+'ZZZZZZZ';
            Pcol^.inf := AGauche(Pcol^.inf,10,'0');
            Pcol^.sup := AGauche(Pcol^.sup,10,'Z');
            Pcol^.typ := 'F';
            Ligne := Format(Formatcol, [Pcol^.cpte, Pcol^.inf, Pcol^.sup,
            Pcol^.typ, 'Collectif Fournisseur']);
            if OkEcritureFichier then writeln(Fichier, Ligne);
            Tcor.PutValue('CR_TYPE', 'SIS');
            Tcor.PutValue('CR_CORRESP', BourreOuTronque(Pcol^.cpte, fbgene));
            Tcor.PutValue('CR_LIBELLE', BourreOuTronque(Pcol^.inf, fbaux));
            Tcor.PutValue('CR_ABREGE', BourreOuTronque(Pcol^.sup, fbaux));
            Tcor.InsertDB(nil, TRUE);
            if Pcol <> nil then Dispose (Pcol);
            New(Pcol);
            Cpte := GetParamSocSecur ('SO_DEFCOLCLI', '4110000000');
            if Copy (Cpte,0,2) <> '40' then Cpte := '4110000000';
            Pcol^.cpte := AGauche(Cpte,10,'0');
            Pcol^.inf := Cercli+'0000000';
            Pcol^.sup := Cercli+'ZZZZZZZ';
            Pcol^.inf := AGauche(Pcol^.inf,10,'0');
            Pcol^.sup := AGauche(Pcol^.sup,10,'Z');
            Pcol^.typ := 'C';
            Ligne := Format(Formatcol, [Pcol^.cpte, Pcol^.inf, Pcol^.sup,
            Pcol^.typ, 'Collectif Client']);
            if OkEcritureFichier then writeln(Fichier, Ligne);
            Tcor.PutValue('CR_TYPE', 'SIS');
            Tcor.PutValue('CR_CORRESP', BourreOuTronque(Pcol^.cpte, fbgene));
            Tcor.PutValue('CR_LIBELLE', BourreOuTronque(Pcol^.inf, fbaux));
            Tcor.PutValue('CR_ABREGE', BourreOuTronque(Pcol^.sup, fbaux));
            if Pcol <> nil then Dispose (Pcol);
            Tcor.InsertDB(nil, TRUE);
            Tcor.Free;

  end
  else
  begin
        while not Q1.EOF do
        begin
            New(Pcol);
            Pcol^.cpte := AGauche(Q1.FindField ('CR_CORRESP').asstring, 10,'0');
            Pcol^.inf := AGauche(Q1.FindField ('CR_LIBELLE').asstring, 10,'0');
            Pcol^.sup := AGauche(Q1.FindField ('CR_ABREGE').asstring, 10,'0');

            if Copy (Pcol^.cpte, 0,2) = '40' then
            begin
               Pcol^.typ := 'F';
               Ligne := Format(Formatcol, [Pcol^.cpte,
               Carfour+copy(Pcol^.inf,2,9),
               Carfour+copy(Pcol^.sup,2,9),
               Pcol^.typ, 'Collectif Fournisseur']);
            end
            else
            begin
               Pcol^.typ := 'C';
               Ligne := Format(Formatcol, [Pcol^.cpte,
               Cercli+copy(Pcol^.inf,2,9),
               Cercli+copy(Pcol^.sup,2,9),
               Pcol^.typ, 'Collectif Client']);
            end;
            writeln(Fichier, Ligne);
            if Pcol <> nil then Dispose (Pcol);
            Q1.next;
        end;
  end;
  Ferme (Q1);

end;

procedure Ecriture_par3(var Fichier : TextFile; Stat : string);
var
ppar3     : ^ar_par3;
Ligne     : string;
begin
    writeln(Fichier, '01');
    new (ppar3);
    ppar3^.cre_cpt := 'N';
    ppar3^.cpt_resul := '1200000000';
    ppar3^.std := '07';
    ppar3^.ana := 'N';
    if (Stat = 'TL') or (Stat = 'ANA') then
        ppar3^.stat := 'O'
    else
        ppar3^.stat := 'N';
    ppar3^.fournisseur := Carfour;
    ppar3^.client := Cercli;
    ppar3^.format := 'P';
    ppar3^.lg_piece := '06';
    ppar3^.anouveau := 'O';
    ppar3^.reper_ecr := 'T';
    ppar3^.monnedit := 'O';
    Ligne := Format(Formatpar3, [ppar3^.siret, ppar3^.ape, ppar3^.tel, ppar3^.std,
    ppar3^.cre_cpt, ppar3^.cpt_resul, ppar3^.ana, ppar3^.stat, ppar3^.fournisseur,
    ppar3^.client, ppar3^.format, ppar3^.lg_piece, ppar3^.info, ppar3^.anouveau,
    ppar3^.tva, ppar3^.devise, ppar3^.v_numc, ppar3^.v_numf,
    ppar3^.reper_ecr, ppar3^.monnedit]);
    writeln(Fichier, Ligne);
    if ppar3 <> nil then Dispose (ppar3);
end;

(*procedure Ecriture_Stat(var Fichier : TextFile);
var
Q1        : Tquery;
Ligne     : string;
begin
   Q1 := OpenSql ('select NT_NATURE,NT_LIBELLE from natcpte where NT_TYPECPTE="E00"', TRUE);

   while not Q1.EOF do
   begin
    Ligne := Format('09%4.4s%25.25s', [Q1.Findfield('NT_NATURE').asstring,Q1.Findfield('NT_LIBELLE').asstring]);
    if Ligne <> '' then
    writeln(Fichier, Ligne);
    Q1.next;
   end;
    Ferme (Q1);
end;
*)

procedure Ecriture_Stat(var Fichier : TextFile; Section : string);
var
Q1        : Tquery;
Ligne     : string;
begin
if Section = 'ANA' then
begin
   Q1 := OpenSql ('select S_SECTION,S_LIBELLE from SECTION', TRUE);

   while not Q1.EOF do
   begin
    Ligne := Format('09%4.4s%25.25s', [Q1.Findfield('S_SECTION').asstring,Q1.Findfield('S_LIBELLE').asstring]);
    if Ligne <> '' then
    writeln(Fichier, Ligne);
    Q1.next;
   end;
    Ferme (Q1);
end
else
begin
   Q1 := OpenSql ('select e_table0 from ecriture where e_table0<>"" group by e_table0', TRUE);

   while not Q1.EOF do
   begin
    Ligne := Format('09%4.4s%-25.25s', [Q1.Findfield('e_table0').asstring,Q1.Findfield('e_table0').asstring]);
    if Ligne <> '' then
    writeln(Fichier, Ligne);
    Q1.next;
   end;
    Ferme (Q1);
end;

end;

procedure Ecriture_par4(var Fichier : TextFile;Trans : TFTransfertcom);
var
ppar4     : ^ar_par4;
Ligne     : string;
Q1        : TQuery;
begin
    new (ppar4);
//    ppar4^.date_encours := '';
    ppar4^.date_encours := FormatDateTime(Traduitdateformat('ddmmyy'),Trans.Dateecr2);
    ppar4^.date_stop := '';
    ppar4^.dex_prev := '12';
    ppar4^.password := '';
    ppar4^.typ := '0';
    ppar4^.res_prev := '000000000000000';
    ppar4^.reopen := '0';
    Q1 := OpenSql ('SELECT EX_ETATCPTA from EXERCICE WHERE EX_EXERCICE="'+Trans.Exo+'"',TRUE);
    if not Q1.Eof then
    begin
      if Q1.FindField ('EX_ETATCPTA').asstring = 'CDE' then
      begin
         ppar4^.cloture := '1';
         // pour envoie de l'exercice clôturé
         ppar4^.date_stop := FormatDateTime(Traduitdateformat('ddmmyy'),Trans.Dateecr2);
      end
      else
         ppar4^.cloture := '0';
    end;
    ppar4^.typ_reouv := '0';
    ferme (Q1);
    Ligne := Format(Formatpar4, [ppar4^.date_encours, ppar4^.date_stop, ppar4^.dex_prev, ppar4^.password,
    ppar4^.typ, ppar4^.res_prev, ppar4^.reopen, ppar4^.cloture, ppar4^.typ_reouv]);
    writeln(Fichier, Ligne);
    if ppar4 <> nil then Dispose (ppar4);
end;


Procedure Ecriture_jrl(var Fichier : TextFile; Where : string) ;
Var
QJ                           : TQuery ;
TJournal                     : ^ar_jour;
SWhere,Ligne,S1              : string;
BEGIN
if Where <> '' then SWhere := Where;
QJ:=OpenSQL('SELECT J_JOURNAL,J_LIBELLE,J_NATUREJAL,J_COMPTEURNORMAL,J_COMPTEURSIMUL,'+
 'J_CONTREPARTIE,J_MODESAISIE,J_COMPTEAUTOMAT,J_COMPTEINTERDIT,J_ABREGE,J_AXE FROM JOURNAL'+ sWhere,TRUE) ;
While Not QJ.Eof Do BEGIN
    new (TJournal);
    TJournal^.journal := copy(QJ.FindField ('J_JOURNAL').asstring,1,2);
    TJournal^.lib := QJ.FindField ('J_LIBELLE').asstring;
    if QJ.FindField ('J_CONTREPARTIE').asstring <> '' then
    TJournal^.cpt_aut := AGauche(QJ.FindField ('J_CONTREPARTIE').asstring,10,'0')
    else
    TJournal^.cpt_aut := '';
    S1:=Copy(TJournal^.cpt_aut,1,3) ;
    If (S1='512') Or (S1='514')Then
       TJournal^.typ := 'T'
    else
    begin
         S1:=Copy(TJournal^.cpt_aut,1,2) ;
         If (S1='53')Then
            TJournal^.typ := 'T'
         else
            TJournal^.typ := 'O';
    end;
    TJournal^.quan := 'N';
    TJournal^.echeance := 'N';
    Ligne := Format(Formatjrl, [TJournal^.journal, TJournal^.lib, TJournal^.cpt_aut,
    TJournal^.sens,TJournal^.typ,TJournal^.quan,TJournal^.echeance,TJournal^.typconf,
    TJournal^.ecart,TJournal^.banque,TJournal^.numcpte,TJournal^.dateimpu,TJournal^.ecart_euro,
    TJournal^.monnaie]);
    writeln(Fichier, Ligne);
    if TJournal <> nil then Dispose (TJournal);
    QJ.NExt ;
END ;
Ferme(QJ) ;

END ;



Procedure FaitSISCOCptTiers(Var FF : TextFile ; Var CritEdt : TCritEdt) ;
Var Q          : TQuery ;
    St         : String ;
    natureaux  : string;
    cpteaux    : string;
    ii         : integer;
BEGIN
TOBTiers := TOB.Create('', nil, -1);
Q:=OpenSQL('SELECT T_AUXILIAIRE,T_LIBELLE, T_TOTDEBP, T_TOTCREP,T_NATUREAUXI FROM TIERS',TRUE) ;

TOBTiers.LoadDetailDB('TIERS', '', '', Q, TRUE, FALSE);
Ferme(Q);

// si changement longueur compte

InitMove(RecordsCount(Q)-1,'') ;
for ii := 0 to TOBTiers.detail.Count - 1 do
BEGIN
  MoveCur(FALSE) ;
  if TOBTiers.detail[ii].GetValue('T_NATUREAUXI') = 'CLI' then natureaux := CerCli
  else
  if TOBTiers.detail[ii].GetValue('T_NATUREAUXI') = 'FOU' then natureaux := Carfour
  else continue;

  if (natureaux = 'F') or (natureaux = 'C') or
     (natureaux = 'f') or (natureaux = 'c') or
     (natureaux = '0') or (natureaux = '9') or (natureaux = '') then
  begin
    Cpteaux := AGauche(TOBTiers.detail[ii].GetValue('T_AUXILIAIRE'),10,'0');
    ControleCompte(Cpteaux);
    if natureaux <> '' then
       Cpteaux := natureaux+Cpteaux;
  if CritEdt.Complement then
  begin
    Cpteaux := BourreLess(Cpteaux, fbAux);
    Cpteaux := AGauche(Cpteaux, 10, ' ');
  end
  else
    Cpteaux := AGauche(Cpteaux, 10, '0');

    St:='C'+Cpteaux+
            AGauche(TOBTiers.detail[ii].GetValue('T_LIBELLE'),25,' ')+ 'NNA' +
            ADroite(StrfMontantSISCO(TOBTiers.detail[ii].GetValue('T_TOTDEBP'),14,CritEdt.Decimale,'',False),14,'0')+
            ADroite(StrfMontantSISCO(TOBTiers.detail[ii].GetValue('T_TOTCREP'),14,CritEdt.Decimale,'',False),14,'0')+
            '      ' ;
    WriteLn(FF,St) ;
  end;
END ;
FiniMove ;
END ;


end.
