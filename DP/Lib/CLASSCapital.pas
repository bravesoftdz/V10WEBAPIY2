{***********UNITE*************************************************
Auteur  ...... : BM
Créé le ...... : 28/04/2003
Modifié le ... : 12/08/2004
Description .. : 
Mots clefs ... : DP;JURI
*****************************************************************}
unit CLASSCapital;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
{$IFDEF EAGLCLIENT}
   MaineAGL, UTOB,
{$ELSE}
   {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
   Controls, sysutils, HCtrls;

/////////////////////////////////////////////////////////////////

type
   TCapital = class
      public
         iNbTitresOuv_c  : integer;
         iNbTitresAugm_c : integer;
         iNbTitresRed_c  : integer;
         iNbTitresClo_c  : integer;

         iNbSoldeAugm_c : integer;
         iNbSoldeRed_c  : integer;

         procedure Init(sGuidPerdos_p : string; iNbTitresClot_p : integer;
                        dExDateDeb_p, dExDateFin_p : TDate;
                        sDomaine_p : string);
         function  IsOk : boolean;
         procedure Ouverture;
         procedure CalculMouvement(var iNbTitresAugm_p, iNbTitresRed_p : integer);
         function  CompareCapital(bAvecMaj_p : boolean) : boolean;

      private
         sGuidPerdos_c   : string;
         dExDateDeb_c    : TDate;
         dExDateFin_c    : TDate;
         sDomaine_c      : string;
end;
/////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////

implementation

/////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 12/08/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TCapital.Init(sGuidPerdos_p : string; iNbTitresClot_p : integer;
                        dExDateDeb_p, dExDateFin_p : TDate;
                        sDomaine_p : string);
begin
   sGuidPerDos_c   := sGuidPerDos_p;
   iNbTitresOuv_c  := 0;
   iNbTitresAugm_c := 0;
   iNbTitresRed_c  := 0;
   iNbTitresClo_c  := iNbTitresClot_p;
   dExDateDeb_c    := dExDateDeb_p;
   dExDateFin_c    := dExDateFin_p;
   sDomaine_c      := sDomaine_p;

//   Ouverture;
   CalculMouvement(iNbTitresAugm_c, iNbTitresRed_c);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 31/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TCapital.IsOk : boolean;
begin
   result := (sGuidPerDos_c <> '') and (dExDateDeb_c <> 0) and (dExDateFin_c <> 0);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 12/08/2004
Modifié le ... :   /  /
Description .. : Nombre de titres à l'ouverture de l'exercice
Mots clefs ... :
*****************************************************************}
procedure TCapital.Ouverture;
var
   QRYReq_l : TQuery;
   iNbTitresOuvn_l : integer;
   iNbTitresOuvp_l : integer;
   sRequete_l : string;
begin
   if not IsOK then Exit;

   iNbTitresOuvn_l := 0;
   iNbTitresOuvp_l := 0;

   // Réduction du nombre de titres avant l'ouverture
   if sDomaine_c = 'JUR' then
      sRequete_l := 'select sum(JMT_NBTITRES) as somme ' +
                    'from JUMVTTITRES ' +
                    'where JMT_GUIDPERDOS = "' + sGuidPerdos_c + '" ' +
                    '  AND JMT_DATE < "' + USDATETIME(dExDateDeb_c) + '" ' +
                    '  AND JMT_SENSOPER = "-" ' +
                    '  AND JMT_VALIDE = "X"'
   else
      sRequete_l := 'select sum(DPM_NBTITRES) as somme ' +
                    'from DPMVTCAP ' +
                    'where DPM_GUIDPER = "' + sGuidPerdos_c + '" ' +
                    '  AND DPM_DATE < "' + USDATETIME(dExDateDeb_c) +'" ' +
                    '  AND DPM_SENS = "-"';

   QRYReq_l := OpenSQL(sRequete_l, TRUE);
   if not QRYReq_l.eof then   // somme <> null
      iNbTitresOuvn_l := (-1) * QRYReq_l.FindField('somme').AsInteger;
   Ferme(QRYReq_l);

   // Augmentation du nombre de titres avant l'ouverture
   if sDomaine_c = 'JUR' then
      sRequete_l := 'select sum(JMT_NBTITRES) as somme ' +
                    'from JUMVTTITRES ' +
                    'where JMT_GUIDPERDOS = "' + sGuidPerdos_c + '" ' +
                    '  AND JMT_DATE < "' + USDATETIME(dExDateDeb_c) + '" ' +
                    '  AND JMT_SENSOPER = "+" ' +
                    '  AND JMT_VALIDE = "X"'
   else
      sRequete_l := 'select sum(DPM_NBTITRES) as somme ' +
                    'from DPMVTCAP ' +
                    'where DPM_GUIDPER = "' + sGuidPerdos_c + '" ' +
                    '  AND DPM_DATE < "' + USDATETIME(dExDateDeb_c) + '" ' +
                    '  AND DPM_SENS = "+"';

   QRYReq_l := OpenSQL(sRequete_l, TRUE);
   if not QRYReq_l.eof then
      iNbTitresOuvp_l := QRYReq_l.FindField('somme').AsInteger;
   Ferme(QRYReq_l);

   iNbTitresOuv_c := iNbTitresOuvn_l + iNbTitresOuvp_l;
//   iNbTitresClo_c := iNbTitresOuv_c;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 21/09/2004
Modifié le ... :   /  /    
Description .. : Dans l'intervalle date de l'exercice
Mots clefs ... : 
*****************************************************************}
procedure TCapital.CalculMouvement(var iNbTitresAugm_p, iNbTitresRed_p : integer);
var
   QRYReq_l : TQuery;
   sRequete_l : string;   
begin
   if not IsOK then Exit;

   // Augmentation du nombre de titres durant l'exercice
   if sDomaine_c = 'JUR' then
      sRequete_l := 'select sum(JMT_NBTITRES) as somme ' +
                    'from JUMVTTITRES ' +
                    'where JMT_GUIDPERDOS = "' + sGuidPerdos_c + '" ' +
                    '  AND JMT_DATE >= "' + USDATETIME(dExDateDeb_c) +'" ' +
                    '  AND JMT_DATE < "' + USDATETIME(dExDateFin_c) + '" ' +
                    '  AND JMT_SENSOPER = "+" ' +
                    '  AND JMT_VALIDE = "X" ' +
                    '  AND JMT_NATUREOP <> "001"'
   else
      sRequete_l := 'select sum(DPM_NBTITRES) as somme ' +
                    'from DPMVTCAP ' +
                    'where DPM_GUIDPER = "' + sGuidPerdos_c + '" ' +
                    '  AND DPM_DATE >= "' + USDATETIME(dExDateDeb_c) +'" ' +
                    '  AND DPM_DATE <= "' + USDATETIME(dExDateFin_c) + '" ' +
                    '  AND DPM_SENS = "+"';

   QRYReq_l := OpenSQL(sRequete_l, TRUE);
   if Not QRYReq_l.eof then
      iNbTitresAugm_p := QRYReq_l.FindField('somme').AsInteger ;
   Ferme(QRYReq_l);

   // Réduction du nombre de titres durant l'exercice
   if sDomaine_c = 'JUR' then
      sRequete_l := 'select sum(JMT_NBTITRES) as somme ' +
                    'from JUMVTTITRES ' +
                    'where JMT_GUIDPERDOS = "' + sGuidPerdos_c + '" ' +
                    '  AND JMT_DATE >= "' + USDATETIME(dExDateDeb_c) + '" ' +
                    '  AND JMT_DATE < "' + USDATETIME(dExDateFin_c) + '" ' +
                    '  AND JMT_SENSOPER = "-" ' +
                    '  AND JMT_VALIDE = "X" ' +
                    '  AND JMT_NATUREOP <> "001"'
   else
      sRequete_l := 'select sum(DPM_NBTITRES) as somme ' +
                    'from DPMVTCAP ' +
                    'where DPM_GUIDPER = "' + sGuidPerdos_c + '" ' +
                    '  AND DPM_DATE >= "' + USDATETIME(dExDateDeb_c) + '" ' +
                    '  AND DPM_DATE <= "' + USDATETIME(dExDateFin_c) + '" ' +
                    '  AND DPM_SENS = "-"';

   QRYReq_l := OpenSQL(sRequete_l, TRUE);
   if not QRYReq_l.eof then
      iNbTitresRed_p := QRYReq_l.FindField('somme').AsInteger;
   Ferme(QRYReq_l);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 19/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TCapital.CompareCapital(bAvecMaj_p : boolean) : boolean;
var
   bChange_l : boolean;
begin
   result := false;
   if not IsOK then Exit;

   bChange_l := false;
   CalculMouvement(iNbSoldeAugm_c, iNbSoldeRed_c);

   if iNbTitresAugm_c <> iNbSoldeAugm_c then
      bChange_l := true;

   if iNbTitresRed_c <> iNbSoldeRed_c then
      bChange_l := true;

   if bAvecMaj_p and bChange_l then
   begin
      iNbTitresAugm_c := iNbSoldeAugm_c;
      iNbTitresRed_c  := iNbSoldeRed_c;
      iNbTitresClo_c := iNbTitresClo_c + iNbTitresAugm_c - iNbTitresRed_c;
   end;
   result := bChange_l;
end;

/////////////////////////////////////////////////////////////////

end.
