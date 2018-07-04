unit UTOMPARAMCA3;

interface
uses
    Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,HDB,db,dbTables,
    UTOM, UTOB, HmsgBox, Hctrls, Ent1, HEnt1,FE_Main,PGIENV,Lia_Commun,HTB97,
    ComCtrls;
procedure MajInfoparamCA3;

type
  TOM_CA3_PARAMETRE = class(TOM)
    procedure OnLoadRecord ; override ;
    procedure RendAccessibleChampca3;
    procedure OnUpdateRecord ; override ;
    private
    procedure ErreurSaisie(mess, page,champ :string) ;
end;

implementation

procedure TOM_CA3_PARAMETRE.OnLoadRecord;
begin
     if (V_PGI_Env <>Nil)  and (V_PGI_Env.ModeFonc='MULTI') then
     begin
            MajInfoparamCA3;
            RendAccessibleChampca3
     end
     else
     SetControlVisible('BMAJFROMDP', FALSE);
end;

procedure TOM_CA3_PARAMETRE.OnUpdateRecord ;
begin
 if GetField ('CP3_NOMFIS') = ''  then
 begin
        ErreurSaisie('Renseigner les Coordonnées des impôts', 'PGeneral','CP3_NOMFIS') ;
        exit;
 end;
 if (GetField ('CP3_JOURDECLA') <= 0) or (GetField ('CP3_JOURDECLA') > 31) then
 begin
        ErreurSaisie('Jour incorrect', 'FE__TabSheet1','CP3_JOURDECLA') ;
        exit;
 end;
 if GetField ('CP3_PERIODICITE') = ''  then
 begin
        ErreurSaisie('Renseigner la périodicité', 'FE__TabSheet1','CP3_PERIODICITE') ;
        exit;
 end;
 if GetField ('CP3_INTRACOM') = ''  then
 begin
        ErreurSaisie('Renseigner le numéro intracommunautaire', 'FE__TabSheet1','CP3_INTRACOM') ;
        exit;
 end;
 if GetField ('CP3_MODEPAIEMENT') = ''  then
 begin
        ErreurSaisie('Renseigner le mode de paiement', 'FE__TabSheet1','CP3_MODEPAIEMENT') ;
        exit;
 end;
 if (GetField ('CP3_REGLEMENTEURO') <> 'F') and (GetField ('CP3_REGLEMENTEURO') <> 'E')then
 begin
        ErreurSaisie('Renseigner la monnaie de réglement', 'FE__TabSheet1','CP3_REGLEMENTEURO') ;
        exit;
 end;
 if GetField ('CP3_DOSSIER') = ''  then
 begin
        ErreurSaisie('Renseigner le numéro du dossier', 'FE__TabSheet1','CP3_DOSSIER') ;
        exit;
 end;
 if GetField ('CP3_CLEDECLA') = ''  then
 begin
        ErreurSaisie('Renseigner la clé', 'FE__TabSheet1','CP3_CLEDECLA') ;
        exit;
 end;

 if GetField ('CP3_REGIME') = ''  then
 begin
        ErreurSaisie('Renseigner le régime', 'FE__TabSheet1','CP3_REGIME') ;
        exit;
 end;

 if GetField ('CP3_INSP') = ''  then
 begin
        ErreurSaisie('Renseigner l''INSP.', 'FE__TabSheet1','CP3_INSP') ;
        exit;
 end;
 if GetField ('CP3_RECETTE') = ''  then
 begin
        ErreurSaisie('Renseigner le champ Recette', 'FE__TabSheet1','CP3_RECETTE') ;
        exit;
 end;

 if GetField ('CP3_CDI') = ''  then
 begin
        ErreurSaisie('Renseigner le champ CDI', 'FE__TabSheet1','CP3_CDI') ;
        exit;
 end;
 if GetField ('CP3_CODEACTIVE') = ''  then
 begin
        ErreurSaisie('Renseigner le code activité', 'FE__TabSheet1','CP3_CODEACTIVE') ;
        exit;
 end;
 if (GetField ('CP3_REDEVABILITE') <> '1') and
    (GetField ('CP3_REDEVABILITE') <> '2') and
    (GetField ('CP3_REDEVABILITE') <> '3') and
    (GetField ('CP3_REDEVABILITE') <> '7') and
    (GetField ('CP3_REDEVABILITE') <> '8') then
 begin
        ErreurSaisie('Saisissez 1,2,3,7 ou 8', 'FE__TabSheet1','CP3_REDEVABILITE') ;
        exit;
 end;

end;


procedure TOM_CA3_PARAMETRE.RendAccessibleChampca3;
begin
          SetControlEnabled ('CP3_JOURDECLA', FALSE);
          SetControlEnabled ('CP3_PERIODICITE', FALSE);
          SetControlEnabled ('CP3_INTRACOM', FALSE);
          SetControlEnabled ('CP3_MODEPAIEMENT', FALSE);
          SetControlEnabled ('CP3_REGLEMENTEURO', FALSE);

          SetControlEnabled ('CP3_DOSSIER', FALSE);
          SetControlEnabled ('CP3_CLEDECLA', FALSE);
          SetControlEnabled ('CP3_REGIME', FALSE);
          SetControlEnabled ('CP3_INSP', FALSE);

          SetControlEnabled ('CP3_NOMFIS', FALSE);
          SetControlEnabled ('CP3_ALRUE1FIS', FALSE);
          SetControlEnabled ('CP3_ALRUE2FIS', FALSE);
          SetControlEnabled ('CP3_ALVILLEFIS', FALSE);
          SetControlEnabled ('CP3_RECETTE', FALSE);
          SetControlEnabled ('CP3_CDI', FALSE);
          SetControlEnabled ('CP3_CODEACTIVE',FALSE);
          SetControlEnabled ('CP3_REDEVABILITE',FALSE)
end;

procedure TOM_CA3_PARAMETRE.ErreurSaisie(mess, page,champ :string) ;
begin
        LastError := 1;
        LastErrorMsg := mess;
        TPageControl(GetControl('Pages')).ActivePage := TTabSheet(GetControl(page));
        SetFocusControl(champ);
end;
procedure MajInfoparamCA3;
var
      Q1,Q2                  : TQuery;
      CP3_NOMFIS             : STRING;
      CP3_ALRUE1FIS          : STRING;
      CP3_ALRUE2FIS          : STRING;
      CP3_ALVILLEFIS         : STRING;
      Periodicite            : STRING;
      Jourech                : integer;
      NOINTRACOMM            : string;
      MODEPAIEFISC           : string;
      REGLEMENTEURO          : string;
      NODOSSINSPEC           : string;
      CLEINSPECT             : string;
      REGIMEINSPECT          : string;
      NOINSPECTION           : string;
      Codeperfis             : integer;
      CDI                    : string;
      RECETTE                : string;
      CODEACTIVITE           : string;
      REDEVABILITE           : string;
      CLESIRET               : string;
begin
     if (V_PGI_Env <>Nil)  and (V_PGI_Env.ModeFonc='MULTI') then
     begin

         Q1 := OpenSql ('SELECT ANN_CODENAF,ANN_CLESIRET FROM ANNUAIRE Where ANN_CODEPER='+ IntToStr (V_PGI_ENV.Codeper), TRUE);
         if not Q1.EOF then
         begin
           CODEACTIVITE := Q1.FindField ('ANN_CODENAF').asstring;
           CLESIRET     := Q1.FindField ('ANN_CLESIRET').asstring;
         end;
         ferme (Q1);

         Q1 := Opensql ('SELECT DFI_PERIODIIMPIND,DFI_JOURDECLA,DFI_NOINTRACOMM,'+
         'DFI_MODEPAIEFISC,DFI_REGLEMENTEURO,DFI_NODOSSINSPEC,DFI_CLEINSPECT,' +
         'DFI_REGIMEINSPECT,DFI_NOINSPECTION,DFI_REDEVABILITE FROM DPFISCAL where DFI_NODP='+IntToStr (V_PGI_ENV.Codeper), TRUE);
         if not Q1.EOF then
         begin
           Periodicite := Q1.FindField ('DFI_PERIODIIMPIND').asstring;
           Jourech := Q1.FindField ('DFI_JOURDECLA').asinteger;
           NOINTRACOMM  := Copy (Q1.FindField ('DFI_NOINTRACOMM').asstring,1,13)+ CLESIRET;
           MODEPAIEFISC := Q1.FindField ('DFI_MODEPAIEFISC').asstring;
           REGLEMENTEURO := Q1.FindField ('DFI_REGLEMENTEURO').asstring;

           NODOSSINSPEC :=  Q1.FindField ('DFI_NODOSSINSPEC').asstring;
           CLEINSPECT := Q1.FindField ('DFI_CLEINSPECT').asstring;
           REGIMEINSPECT := Q1.FindField ('DFI_REGIMEINSPECT').asstring;
           NOINSPECTION := Q1.FindField ('DFI_NOINSPECTION').asstring;
           REDEVABILITE := Q1.FindField ('DFI_REDEVABILITE').asstring;
         end;
         ferme (Q1);

         if not ExisteSQL('SELECT CP3_DATEENCOURS FROM CA3_PARAMETRE') then
           ExecuteSql ('Insert into CA3_PARAMETRE (CP3_CODEPER,CP3_PERIODICITE,CP3_JOURDECLA,CP3_INTRACOM,'+
                      'CP3_MODEPAIEMENT,CP3_REGLEMENTEURO,CP3_DOSSIER,CP3_CLEDECLA,CP3_REGIME,CP3_INSP,CP3_REDEVABILITE,CP3_CODEACTIVE)' +
                      'Values ('+IntToStr (V_PGI_ENV.Codeper)+',"'+Periodicite+'",'+ IntToStr(Jourech)+',"'+ NOINTRACOMM +
                      '","'+ MODEPAIEFISC+'","' + REGLEMENTEURO+'","'+ NODOSSINSPEC+'","'+ CLEINSPECT+'","'+ REGIMEINSPECT+
                      '","'+ NOINSPECTION+'","'+ REDEVABILITE+
                      '","'+ CODEACTIVITE+'")')
         else
           ExecuteSql ('Update CA3_PARAMETRE SET CP3_PERIODICITE="'+ Periodicite
                      + '", CP3_JOURDECLA='+ IntToStr (Jourech) + ', CP3_INTRACOM="'+ NOINTRACOMM
                      + '", CP3_MODEPAIEMENT="'+ MODEPAIEFISC + '", CP3_REGLEMENTEURO="'+ REGLEMENTEURO
                      + '", CP3_DOSSIER="'+  NODOSSINSPEC+ '", CP3_CLEDECLA="'+ CLEINSPECT+'", CP3_REGIME="'+ REGIMEINSPECT
                      + '", CP3_INSP="'+ NOINSPECTION
                      + '", CP3_REDEVABILITE="'+ REDEVABILITE
                      + '", CP3_CODEACTIVE="'+ CODEACTIVITE+'" WHERE CP3_CODEPER='+IntToStr (V_PGI_ENV.Codeper));


        Q1 := OpenSql ('SELECT ANL_CODEPER,ANL_NOADHESION,ANL_ORGANISME from ANNULIEN Where ANL_CODEPERDOS='+IntToStr (V_PGI_ENV.Codeper) + ' AND ANL_FONCTION="FIS" AND ANL_TYPEPER="RDI"', TRUE);
        if Q1.EOF then begin  Ferme (Q1); exit; end
        else
        begin
            Codeperfis := Q1.FindField ('ANL_CODEPER').asinteger;
            RECETTE :=  Q1.Findfield ('ANL_NOADHESION').asstring;
            CDI :=  copy (Q1.Findfield ('ANL_ORGANISME').asstring,0,3);
            Ferme (Q1);
        end;
        Q1 := OpenSql ('SELECT ANN_NOMPER,ANN_ALRUE1,ANN_ALRUE2,ANN_ALCP,ANN_ALVILLE FROM ANNUAIRE Where ANN_CODEPER='+ IntToStr (Codeperfis), TRUE);
        if not Q1.EOF then
        begin
             CP3_NOMFIS := Q1.Findfield ('ANN_NOMPER').asstring;
             CP3_ALRUE1FIS := Q1.Findfield ('ANN_ALRUE1').asstring;
             CP3_ALRUE2FIS :=  Q1.Findfield ('ANN_ALRUE2').asstring;
             CP3_ALVILLEFIS := Q1.Findfield ('ANN_ALCP').asstring+ ' '+ Q1.Findfield ('ANN_ALVILLE').asstring;
             if ExisteSQL('SELECT CP3_DATEENCOURS FROM CA3_PARAMETRE') then
                  ExecuteSql ('Update CA3_PARAMETRE SET CP3_NOMFIS="'+ CP3_NOMFIS
                             + '", CP3_ALRUE1FIS="'+ CP3_ALRUE1FIS
                             + '", CP3_ALRUE2FIS="'+ CP3_ALRUE2FIS+ '", CP3_ALVILLEFIS="'+ CP3_ALVILLEFIS
                             + '", CP3_RECETTE="'+ RECETTE+ '", CP3_CDI="'+ CDI +
                             '" WHERE CP3_CODEPER='+IntToStr (V_PGI_ENV.Codeper));
        end;
        Ferme (Q1);
     end;
end;

Initialization
registerclasses([TOM_CA3_PARAMETRE]);

end.
