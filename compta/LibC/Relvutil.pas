unit RelvUtil;

interface

uses
    Windows,
    Messages,
    SysUtils,
    Classes,
    Graphics,
    Controls,
    Forms,
    Hctrls,
    Ent1,
    Hent1,
    uTOB,
    UtilSais, // pour MajSoldesEcritureTOB
{$IFNDEF EAGLCLIENT}
    DB,
    {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
    DBGrids,
    HDB,
    DBCtrls,
    utilSoc, // pour MarquerPublifi
{$ENDIF}
    SaisUtil,
    SaisComm,
    ValPerio,
    HStatus,
    Echeance,
    Hcompte;

type
    EC_RELF = record
        General, Auxiliaire:String17;
        DebitD, CreditD:Double;
        DebitP, CreditP:Double;
//        DebitE, CreditE:Double;
        ModePaie, RegimeTva:String3;
        Echeance:TDateTime;
        RefRel:String35;
        CodeL:String4;
        DateMin, DateMax:TDateTime;
        TabTvaEnc:array[1..5] of Double;
    end;

procedure GenerePieceReleve(REL:EC_RELF;M:RMVT;DEV:RDEVISE);

implementation

uses
  {$IFDEF MODENT1}
  CPProcGen,
  {$ENDIF MODENT1}
  ULibecriture;


procedure InitFrom(TEcr : TOB ; T_Libelle:string;REL:EC_RELF;M:RMVT;DEV:RDEVISE;i:integer);
begin
    TEcr.InitValeurs;
    TEcr.PutValue('E_EXERCICE', M.Exo);
    TEcr.PutValue('E_JOURNAL', M.Jal);
    TEcr.PutValue('E_DATECOMPTABLE', M.DateC);
    TEcr.PutValue('E_NUMEROPIECE', M.Num);
{$IFNDEF SPEC302}
    TEcr.PutValue('E_PERIODE', GetPeriode(M.DateC));
    TEcr.PutValue('E_SEMAINE', NumSemaine(M.DateC));
{$ENDIF}
    TEcr.PutValue('E_GENERAL', REL.General);
    TEcr.PutValue('E_AUXILIAIRE', REL.Auxiliaire);
    TEcr.PutValue('E_NATUREPIECE', M.Nature);
    TEcr.PutValue('E_QUALIFPIECE', M.Simul);
    TEcr.PutValue('E_ETABLISSEMENT', M.Etabl);
    TEcr.PutValue('E_ECRANOUVEAU', 'N');
    TEcr.PutValue('E_MODEPAIE', REL.ModePaie);
    TEcr.PutValue('E_DATEECHEANCE', REL.Echeance);
    TEcr.PutValue('E_DEVISE', M.CodeD);
    TEcr.PutValue('E_TAUXDEV', M.TauxD);
    TEcr.PutValue('E_DATETAUXDEV', M.DateTaux);
    TEcr.PutValue('E_NUMECHE', 1);
    TEcr.PutValue('E_REFRELEVE', REL.RefRel);
    TEcr.PutValue('E_REFINTERNE', REL.RefRel);
    TEcr.PutValue('E_ANA', '-');
    TEcr.PutValue('E_ECHE', 'X');
    TEcr.PutValue('E_REGIMETVA', REL.RegimeTva);
    TEcr.PutValue('E_ENCAISSEMENT', '');
    TEcr.PutValue('E_TVAENCAISSEMENT', '-');

    if M.Valide then TEcr.PutValue('E_VALIDE', 'X')
    else TEcr.PutValue('E_VALIDE', '-');

{MAJ depuis Tiers}
    if T_Libelle <> '' then TEcr.PutValue('E_LIBELLE', T_Libelle);
{Lettrage total première ligne}
    if i = 1 then
    begin
        TEcr.PutValue('E_LETTRAGE', REL.CodeL);
        TEcr.PutValue('E_ETATLETTRAGE', 'TL');
        TEcr.PutValue('E_DATEPAQUETMIN', REL.DateMin);
        TEcr.PutValue('E_DATEPAQUETMAX', REL.DateMax);
        if DEV.Code <> V_PGI.DevisePivot then
        begin
            TEcr.PutValue('E_LETTRAGEDEV', 'X');
        end
        else
        begin
            TEcr.PutValue('E_LETTRAGEDEV', '-');
        end;
        TEcr.PutValue('E_ETAT', '---0RM----');
        TEcr.PutValue('E_FLAGECR', 'RDE'); {Relevé Destination}
    end
    else
    begin
        TEcr.PutValue('E_ETATLETTRAGE', 'AL');
        TEcr.PutValue('E_FLAGECR', 'REC'); {Relevé Ecriture}
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 01/04/2004
Modifié le ... :   /  /    
Description .. : - 01/04/2004 - LG - passgae en eAGL ( modif suite a la 
Suite ........ : reecriture deAboUtil )
Mots clefs ... : 
*****************************************************************}
procedure GenerePieceReleve(REL:EC_RELF;M:RMVT;DEV:RDEVISE);
var
    TEcr:TOB;
    i, k:integer;
    Q:TQuery;
    Coll:string;
    DD, CD, DP, CP, Coef:Double;
    TPiece:TOB;
    T_Libelle : string ;
begin
    DD := REL.DebitD;
    CD := REL.CreditD;
    DP := REL.DebitP;
    CP := REL.CreditP;
//    DE := REL.DebitE;
//    CE := REL.CreditE;
//    Coef := 1.0;
    Q := OpenSQL('Select T_LIBELLE from TIERS Where T_AUXILIAIRE="' + REL.Auxiliaire + '"', True);
    T_Libelle := Q.FindField('T_LIBELLE').AsString ;
    Ferme(Q) ;

    TPiece := TOB.Create('vPiece', nil, -1);


    for i := 1 to 2 do
    begin

        TEcr := TOB.Create('ECRITURE', TPiece, -1);
        CPutDefautEcr(TEcr) ;
        InitFrom(TEcr, T_Libelle, REL, M, DEV, i);
        TEcr.PutValue('E_NUMLIGNE', i);
        TEcr.PutValue('E_TYPEMVT', 'TTC');
        if i = 1 then
        begin
            Coef := -1.0;
            if VH^.MontantNegatif then
            begin
                TEcr.PutValue('E_DEBITDEV', -DD);
                TEcr.PutValue('E_CREDITDEV', -CD);
                TEcr.PutValue('E_DEBIT', -DP);
                TEcr.PutValue('E_CREDIT', -CP);
                TEcr.PutValue('E_COUVERTUREDEV', -DD - CD);
                TEcr.PutValue('E_COUVERTURE', -DP - CP);
            end
            else
            begin
                TEcr.PutValue('E_DEBIT', CP);
                TEcr.PutValue('E_CREDIT', DP);
                TEcr.PutValue('E_COUVERTURE', DP + CP);
                TEcr.PutValue('E_DEBITDEV', CD);
                TEcr.PutValue('E_CREDITDEV', DD);
                TEcr.PutValue('E_COUVERTUREDEV', DD + CD);
            end;
        end
        else
        begin
            Coef := +1.0;
            TEcr.PutValue('E_DEBITDEV', DD);
            TEcr.PutValue('E_CREDITDEV', CD);
            TEcr.PutValue('E_DEBIT', DP);
            TEcr.PutValue('E_CREDIT', CP);
        end;
        if VH^.OuiTvaEnc then
        begin
            Coll := REL.General;
            if EstCollFact(Coll) then
            begin
                for k := 1 to 4 do
                    TEcr.PutValue('E_ECHEENC' + IntToStr(k), Arrondi(Coef * REL.TabTvaEnc[k], V_PGI.OkDecV));
                TEcr.PutValue('E_ECHEDEBIT', Arrondi(Coef * REL.TabTvaEnc[5], V_PGI.OkDecV));
                TEcr.PutValue('E_EMETTEURTVA', 'X');
            end;
        end;
        TEcr.PutValue('E_ENCAISSEMENT', SensEnc(TEcr.GetValue('E_DEBIT'),
            TEcr.GetValue('E_CREDIT')));

        CSetCotation(TEcr);
        TEcr.PutValue('E_CODEACCEPT', MPTOACC( TEcr.GetValue('E_MODEPAIE') ) ) ;

        TEcr.InsertDB(nil);

    end;

 MajSoldesEcritureTOB(TPiece, True);
 ADevalider(M.Jal, M.DateC);

{$IFNDEF EAGLCLIENT}
    MarquerPublifi(True);
{$ENDIF}

 CPStatutDossier ;

end;

end.

