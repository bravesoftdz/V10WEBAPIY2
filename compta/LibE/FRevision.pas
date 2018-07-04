unit FRevision;

interface

uses
    classes,
    Forms,
    Sysutils,
    Ent1,HEnt1,
    UTob,
    {$IFNDEF EAGLCLIENT}
    dbtables,
    {$Else}
    UtileAGL,
    {$ENDIF}
     dialogs, HCtrls, StdCtrls,Controls;


 (* Windows,
  Messages,
  Graphics,
  Grids,
{$IFNDEF EAGLCLIENT}
  dbtables,
{$ENDIF}
  HCtrls, StdCtrls;*)

type
  TF_Revision = class(TForm)
    Memo1: TMemo;
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

Procedure LancerRevision;
Function GetInfoCRE : Variant;

const
   NumEmprunt  = 0;
   LibEmp      = 1;
   NumCompte   = 2;
   LibCompte   = 3;
   DateContrat = 4;
   Duree       = 5;
   Capital     = 6;
   CapNouvEmp  = 7;

   AmoExe      = 8;
   Moins1an    = 9;
   De1a5an     = 10;
   Plus5ans    = 11;

   IntExe      = 12;
   IntCourus1  = 13;
   IntCourus2  = 14;
   IntApresExe = 15;

   AssExe      = 16;
   AssCourus1  = 17;
   AssCourus2  = 18;
   AssApresExe = 19;

   cNbCols     = 20;


implementation

{$R *.DFM}



Procedure LancerRevision;
var
   FRevision : TF_Revision;
begin
   Application.CreateForm(TF_Revision,FRevision);
   try
      FRevision.Memo1.Text := GetInfoCRE;
      FRevision.ShowModal;
   finally
      FRevision.Release;
   end;
end;


Function GetInfoCRE : Variant;
var
   x:string;
   lObQuery       : TQuery;
   lInLastEmpCode : Integer;
   lDtDateEch     : TDateTime;
   lRdAmoEch      : Double;
   lRdIntEch      : Double;
   lRdAssEch      : Double;
   lRdIntEchPrec  : Double;
   lRdAssEchPrec  : Double;
   lDtDate1       : TDateTime;
   lDtDate2       : TDateTime;
   lDtDate3       : TDateTime;
   lDtDate4       : TDateTime;
   lDtDateEchPrec : TDateTime;
   lRdProrata     : Double;
   lRdTotalExe    : Double;
   lRdTotalMoins1 : Double;
   lRdTotal1a5    : Double;
   lRdTotalPlus5  : Double;
   lRdIntExe      : Double;
   lRdIntApresExe : Double;
   lRdAssExe      : Double;
   lRdIntCourus1  : Double;
   lRdIntCourus2  : Double;
   lRdAssCourus1  : Double;
   lRdAssCourus2  : Double;
   lRdIntCCA1     : Double;
   lRdIntCCA2     : Double;
   lRdAssCCA1     : Double;
   lRdAssCCA2     : Double;
   lRdAssApresExe : Double;
   lInAn,
   lInMois ,
   lInJour,lInJour2 : Word;
   lBoRupture     : Boolean;
   lStSQL         : String;
   lStValeurs    : String;
   lStNomsChamps : String;
   lObResult     : TStringList;
   lBoManquant   : Boolean;
const
   cSeparateur = ';';
   cRetourChariot = chr(13);

   //
   Procedure Stocker(pStNomChamps : String; pStValeur : String);
   begin
      pStNomChamps := Trim(Uppercase(pStNomChamps));

      if lStNomsChamps <> '' then
      begin
         lStNomsChamps := lStNomsChamps + cSeparateur;
         lStValeurs := lStValeurs + cSeparateur;
      end
      else
         if lStValeurs <> '' then
            lStValeurs := lStValeurs + cRetourChariot;

      lStNomsChamps := lStNomsChamps + pStNomChamps;
      lStValeurs := lStValeurs + pStValeur;
   end;

begin

   result := '#Erreur : Non connecté' ;

   if Not V_PGI.OKOuvert then Exit ;

   lObQuery := Nil;
   lObResult := Nil;
   lStValeurs := '';

   //
   // Calcul des dates
   //
   // Date 1 : début d'exercice
    lDtDate1 := VH^.Encours.Deb;
   decodedate(lDtDate1, lInAn, lInMois, lInJour);
   // Date 2 : début d'exercice suivant
   try
    //Djamel : date début exercice N+1 = date début exercice N + 12 mois
    //  lDtDate2 := EncodeDate(lInAn + 1, lInMois, lInJour);
    (*CAD : date début exercice N+1 = date fin exercice N + 1 jour*)
    lDtDate2 := VH^.Encours.Fin;
    lDtDate2 := lDtDate2 + 1;
    decodedate(lDtDate2, lInAn, lInMois, lInJour2);
   except
      lDtDate2 := EncodeDate(lInAn + 1, lInMois, 28);
   end;
   // Date 3 : Exercice suivant
   try
   //   lDtDate3 := EncodeDate(lInAn + 2, lInMois, lInJour);
      lDtDate3 := EncodeDate(lInAn + 1, lInMois, lInJour);
   except
   //   lDtDate3 := EncodeDate(lInAn + 2, lInMois, 28);
      lDtDate3 := EncodeDate(lInAn + 1, lInMois, 28);
   end;
   // Date 4 : + d'un an et - de 5 ans
   try
  //    lDtDate4 := EncodeDate(lInAn + 6, lInMois, lInJour);
      lDtDate4 := EncodeDate(lInAn + 5, lInMois, lInJour);
   except
  //    lDtDate4 := EncodeDate(lInAn + 6, lInMois, 28);
      lDtDate4 := EncodeDate(lInAn + 5, lInMois, 28);
   end;

   try
      lStSQL := ' Select ECH_DATE, ECH_AMORTISSEMENT, ECH_INTERET,' +
                '  ECH_ASSURANCE, ECH_VERSEMENT, ECH_DETTE,' +
                '  ECH_DATEEXERCICE, EMP_CODEEMPRUNT, EMP_LIBEMPRUNT,' +
                '  EMP_NUMCOMPTE, EMP_DATEDEBUT, EMP_DATECONTRAT,' +
                '  EMP_CAPITAL, EMP_DUREE, EMP_TOTINT,' +
                '  EMP_TOTASS, EMP_TOTAMORT, EMP_TOTVERS,' +
                '  EMP_NBECHTOT, G_GENERAL, G_LIBELLE' +
                ' From FECHEANCE, FEMPRUNT' +
                ' Left Join GENERAUX On G_GENERAL = EMP_NUMCOMPTE' +
                ' Where ECH_CODEEMPRUNT = EMP_CODEEMPRUNT' +
                  ' And EMP_CODEEMPRUNT > -1' +
                ' Order By EMP_CODEEMPRUNT, ECH_DATE';

      lObQuery := OpenSQL(lStSQL, True);

      lRdTotalExe    := 0;
      lRdTotalMoins1 := 0;
      lRdTotal1a5    := 0;
      lRdTotalPlus5  := 0;

      lRdIntExe      := 0;
      lRdIntCourus1  := 0;
      lRdIntCourus2  := 0;
      lRdIntApresExe := 0;

      lRdAssExe      := 0;
      lRdAssCourus1  := 0;
      lRdAssCourus2  := 0;
      lRdAssApresExe := 0;

      lRdIntCCA1  := 0;
      lRdIntCCA2  := 0;
      lRdAssCCA1  := 0;
      lRdAssCCA2  := 0;

      lDtDateEchPrec := -1;

      lInLastEmpCode := -1000;

      //
      //  Boucle
      //
      While not lObQuery.Eof do
      begin
         // Entête de bande détail
         if lObQuery.FindField('EMP_CODEEMPRUNT').AsInteger <> lInLastEmpCode then
         begin
            lStNomsChamps := '';

            lRdTotalExe    := 0;
            lRdTotalMoins1 := 0;
            lRdTotal1a5    := 0;
            lRdTotalPlus5  := 0;

            lRdIntExe      := 0;
            lRdIntCourus1  := 0;
            lRdIntCourus2  := 0;
            lRdIntApresExe := 0;

            lRdAssExe      := 0;
            lRdAssCourus1  := 0;
            lRdAssCourus2  := 0;
            lRdAssApresExe := 0;

            lRdIntCCA1 := 0;
            lRdIntCCA2 := 0;
            lRdAssCCA1 := 0;
            lRdAssCCA2 := 0;

            lDtDateEchPrec := -1;
         //   lDtDateEchPrec := lObQuery.FindField('EMP_DATECONTRAT').AsDateTime;
//Correction CAD du 07/12/04 : la date de la première échéance ne doit pas être celle de la date du contrat !!!
            lBoManquant := lObQuery.Findfield('EMP_DATECONTRAT').AsDateTime < lObQuery.FindField('EMP_DATEDEBUT').AsDateTime;

            Stocker('NumEmprunt',  lObQuery.FindField('EMP_CODEEMPRUNT').AsString);
            Stocker('LibEmp',      lObQuery.FindField('EMP_LIBEMPRUNT').AsString);
            Stocker('NumCompte',   lObQuery.FindField('EMP_NUMCOMPTE').AsString);
            Stocker('LibCompte',   lObQuery.FindField('G_LIBELLE').AsString);
            Stocker('DateContrat', lObQuery.FindField('EMP_DATECONTRAT').AsString);
            Stocker('Duree',       lObQuery.FindField('EMP_DUREE').AsString);
            Stocker('Capital',     lObQuery.FindField('EMP_CAPITAL').AsString);
            if (lObQuery.FindField('EMP_DATECONTRAT').AsDateTime >= lDtDate1) and
               (lObQuery.FindField('EMP_DATECONTRAT').AsDateTime <  lDtDate2) then
               Stocker('CapNouvEmp', lObQuery.FindField('EMP_CAPITAL').AsString)
            else
               Stocker('CapNouvEmp', '0');
         end;

         lDtDateEch := lObQuery.FindField('ECH_DATE').AsDateTime;
         lRdAmoEch  := lObQuery.FindField('ECH_AMORTISSEMENT').AsFloat;
         lRdIntEch  := lObQuery.FindField('ECH_INTERET').AsFloat;
         lRdAssEch  := lObQuery.FindField('ECH_ASSURANCE').AsFloat;

         // Calcul des totaux
         if (lDtDateEch >= lDtDate1) and (lDtDateEch < lDtDate2) then
         begin
            lRdTotalExe := lRdTotalExe + lRdAmoEch;
            lRdIntExe   := lRdIntExe + lRdIntEch;
            lRdAssExe := lRdAssExe + lRdAssEch;
         end
         else
            if (lDtDateEch >= lDtDate2) and (lDtDateEch < lDtDate3) then
            begin
               lRdTotalMoins1 := lRdTotalMoins1 + lRdAmoEch;
            end
            else
               if (lDtDateEch >= lDtDate3) and (lDtDateEch < lDtDate4) then
                  lRdTotal1a5 := lRdTotal1a5 + lRdAmoEch
               else
                  if lDtDateEch >= lDtDate4 then
                     lRdTotalPlus5 := lRdTotalPlus5 + lRdAmoEch;

         if lDtDateEch >= lDtDate2 then
         begin
            lRdIntApresExe := lRdIntApresExe + lRdIntEch;
            lRdAssApresExe := lRdAssApresExe + lRdAssEch;
         end;

         //
         // Courus N - 1
         //
         // si la date de fin d'exercice N-1 est entre l'échéance précédente
         // et l'échéance courante
         if lDtDateEchPrec>-1 then
         begin
             if (lDtDate1 - 1 > lDtDateEchPrec) and (lDtDate1 - 1 <= lDtDateEch) then
             begin
                if lBoManquant then
                begin
                   // Intérets manquants
                   lRdProrata := (lDtDate1 - 1 - lDtDateEchPrec) / (lDtDateEch - lDtDateEchPrec);
                   lRdIntCourus1 := lRdProrata * lRdIntEch;
                   lRdAssCourus1 := lRdProrata * lRdAssEch;
                end
                else
                begin
                   lRdProrata := (lDtDateEch - lDtDate1 + 1) / (lDtDateEch - lDtDateEchPrec);
                   lRdIntCCA1 := lRdProrata * lRdIntEchPrec;
                   lRdAssCCA1 := lRdProrata * lRdAssEchPrec;
                end;
             end;
              //
             // Courus N
             //
             // si la date de fin d'exercice est entre l'échéance précédente
             // et l'échéance courante
             if (lDtDate2 - 1 > lDtDateEchPrec) and (lDtDate2 - 1 <= lDtDateEch) then
             begin
                if lBoManquant then
                begin
                   // Intérets manquants
                   lRdProrata := (lDtDate2 - 1 - lDtDateEchPrec) / (lDtDateEch - lDtDateEchPrec);
                   lRdIntCourus2 := lRdProrata * lRdIntEch;
                   lRdAssCourus2 := lRdProrata * lRdAssEch;
                end
                else
                begin
                   lRdProrata := (lDtDateEch - lDtDate2 + 1) / (lDtDateEch - lDtDateEchPrec);
                   lRdIntCCA2 := lRdProrata * lRdIntEchPrec;
                   lRdAssCCA2 := lRdProrata * lRdAssEchPrec;
                end;
             end;
         end
      //******************************************************
         else
         begin
         //
         // Courus N
         //
         // 1. Si la date de mise à disposition des fonds est antérieure à la date de première échéance et
         // 2  si la date de mise à disposition des fonds se situe dans l'exercice et
         // 3  si la date de la première échéance est postérieure à l'exercice.
            if lBoManquant then
            begin
                lDtDateEchPrec := lObQuery.FindField('EMP_DATECONTRAT').AsDateTime;
             // si la date de fin d'exercice est entre l'échéance précédente
             // et l'échéance courante
                 if (lDtDate1 - 1 < lDtDateEchPrec) and (lDtDate2 - 1 > lDtDateEchPrec) and(lDtDate2 - 1 <= lDtDateEch) then
                 begin
                                // Intérets manquants
                   lRdProrata := (lDtDate2 - 1 - lDtDateEchPrec) / (lDtDateEch - lDtDateEchPrec);
                   lRdIntCourus2 := lRdProrata * lRdIntEch;
                   lRdAssCourus2 := lRdProrata * lRdAssEch;
                end;
            end;
         end;
      //******************************************************
         // Echéance suivante
         lDtDateEchPrec := lDtDateEch;
         lRdIntEchPrec  := lRdIntEch;
         lRdAssEchPrec  := lRdAssEch;
         lInLastEmpCode := lObQuery.FindField('EMP_CODEEMPRUNT').AsInteger;
         lObQuery.Next;

         // pied de bande détail
         if lObQuery.Eof then
            lBoRupture := True
         else
            lBoRupture := (lInLastEmpCode <> lObQuery.FindField('EMP_CODEEMPRUNT').AsInteger);
         if lBoRupture then
         begin
            Stocker('AmoExe',   FloatTostr(lRdTotalExe));
            Stocker('Moins1an', FloatTostr(lRdTotalMoins1));
            Stocker('De1a5an',  FloatTostr(lRdTotal1a5));
            Stocker('Plus5ans', FloatTostr(lRdTotalPlus5));

            Stocker('IntExe',      FloatTostr(lRdIntExe));
            Stocker('IntCourus1',  FloatTostr(lRdIntCourus1));
            Stocker('IntCCA1',     FloatTostr(lRdIntCCA1));
            Stocker('IntCourus2',  FloatTostr(lRdIntCourus2));
            Stocker('IntCCA2',     FloatTostr(lRdIntCCA2));
            Stocker('IntApresExe', FloatTostr(lRdIntApresExe));

            Stocker('AssExe',      FloatTostr(lRdAssExe));
            Stocker('AssCourus1',  FloatTostr(lRdAssCourus1));
            Stocker('AssCCA1',     FloatTostr(lRdAssCCA1));
            Stocker('AssCourus2',  FloatTostr(lRdAssCourus2));
            Stocker('AssCCA2',     FloatTostr(lRdAssCCA2));
            Stocker('AssApresExe', FloatTostr(lRdAssApresExe));
         end;
      end;

      // Résultat
      Result := lStNomsChamps + cRetourChariot + lStValeurs;
      lObResult := TStringList.Create;
      try
         lObResult.Text := result;
         lObResult.SaveToFile('C:\CRE.CSV');
      except
      end;

   finally
      if Assigned(lObQuery) then Ferme(lObQuery);
      if Assigned(lObResult) then lObResult.Free;;
   end;

end;

end.

