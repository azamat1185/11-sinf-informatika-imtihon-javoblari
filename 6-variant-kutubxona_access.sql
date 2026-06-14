/* ================================================================
   KUTUBXONA BOSHQARUV TIZIMI — Microsoft Access SQL
   Library Management System for Microsoft Access (.accdb)
   ================================================================
   Muallif: 1-Topshiriq yechimi
   Sana: 2026
   ================================================================

   QANDAY ISHLATISH:
   1. Microsoft Access > Blank Database yarating
   2. Database Tools > Visual Basic (VBE) oching
   3. Insert > Module qo'shing
   4. kutubxona_setup.vba faylidagi kodni yapishtirib Execute qiling
   YOKI:
   1. Access > Create > Query Design oching
   2. SQL View ga o'ting
   3. Quyidagi CREATE TABLE blokini birma-bir bajaring

   ================================================================ */


/* ================================================================
   1. JADVALLAR YARATISH / CREATE TABLES
   ================================================================ */

/* -- Kitoblar jadvali */
CREATE TABLE Kitoblar (
    ID        AUTOINCREMENT   CONSTRAINT PK_Kitoblar PRIMARY KEY,
    Nomi      TEXT(255)       NOT NULL,
    Muallif   TEXT(255)       NOT NULL,
    ISBN      TEXT(20)        NOT NULL,
    Soni      INTEGER         NOT NULL DEFAULT 1,
    CONSTRAINT UQ_ISBN UNIQUE (ISBN)
);

/* -- Azolar jadvali */
CREATE TABLE Azolar (
    ID           AUTOINCREMENT   CONSTRAINT PK_Azolar PRIMARY KEY,
    Ism          TEXT(100)       NOT NULL,
    Familiya     TEXT(100)       NOT NULL,
    Azo_raqami   TEXT(20)        NOT NULL,
    CONSTRAINT UQ_AzoRaqami UNIQUE (Azo_raqami)
);

/* -- Ijara jadvali */
CREATE TABLE Ijara (
    ID                INTEGER       CONSTRAINT PK_Ijara PRIMARY KEY,
    Kitob_ID          INTEGER       NOT NULL,
    Azo_ID            INTEGER       NOT NULL,
    Olingan_sana      DATETIME      NOT NULL,
    Qaytarish_sanasi  DATETIME      NOT NULL,
    CONSTRAINT FK_Ijara_Kitob FOREIGN KEY (Kitob_ID) REFERENCES Kitoblar(ID),
    CONSTRAINT FK_Ijara_Azo   FOREIGN KEY (Azo_ID)   REFERENCES Azolar(ID)
);

/* -- Javobgarlik jadvali */
CREATE TABLE Javobgarlik (
    ID                              INTEGER     CONSTRAINT PK_Javobgarlik PRIMARY KEY,
    Ijara_ID                        INTEGER     NOT NULL,
    Kechiktirilganlik_uchun_jarima  CURRENCY    NOT NULL DEFAULT 0,
    CONSTRAINT FK_Javob_Ijara FOREIGN KEY (Ijara_ID) REFERENCES Ijara(ID)
);


/* ================================================================
   2. MA'LUMOT QO'SHISH / INSERT DATA
   ================================================================ */

/* -- Kitoblar */
INSERT INTO Kitoblar (Nomi, Muallif, ISBN, Soni)
VALUES ("O'tkan kunlar", "Abdulla Qodiriy", "978-9943-01-001-1", 5);

INSERT INTO Kitoblar (Nomi, Muallif, ISBN, Soni)
VALUES ("Mehrobdan chayon", "Abdulla Qodiriy", "978-9943-01-002-2", 3);

INSERT INTO Kitoblar (Nomi, Muallif, ISBN, Soni)
VALUES ("Sarob", "Abdulla Qahhor", "978-9943-01-003-3", 4);

INSERT INTO Kitoblar (Nomi, Muallif, ISBN, Soni)
VALUES ("Python dasturlash", "Mark Lutz", "978-0-596-15806-4", 2);

INSERT INTO Kitoblar (Nomi, Muallif, ISBN, Soni)
VALUES ("Clean Code", "Robert C. Martin", "978-0-13-235088-4", 1);

/* -- Azolar */
INSERT INTO Azolar (Ism, Familiya, Azo_raqami)
VALUES ("Akbar", "Toshmatov", "AZO-0001");

INSERT INTO Azolar (Ism, Familiya, Azo_raqami)
VALUES ("Barno", "Xasanova", "AZO-0002");

INSERT INTO Azolar (Ism, Familiya, Azo_raqami)
VALUES ("Jasur", "Normatov", "AZO-0003");

INSERT INTO Azolar (Ism, Familiya, Azo_raqami)
VALUES ("Dilnoza", "Yusupova", "AZO-0004");

INSERT INTO Azolar (Ism, Familiya, Azo_raqami)
VALUES ("Sardor", "Mirzayev", "AZO-0005");

/* -- Ijara (namunaviy, muddati o'tgan va joriy) */
INSERT INTO Ijara (ID, Kitob_ID, Azo_ID, Olingan_sana, Qaytarish_sanasi)
VALUES (1, 1, 1, #5/1/2025#, #5/15/2025#);

INSERT INTO Ijara (ID, Kitob_ID, Azo_ID, Olingan_sana, Qaytarish_sanasi)
VALUES (2, 2, 2, #5/10/2025#, #5/24/2025#);

INSERT INTO Ijara (ID, Kitob_ID, Azo_ID, Olingan_sana, Qaytarish_sanasi)
VALUES (3, 3, 3, #6/1/2026#, #6/15/2026#);

INSERT INTO Ijara (ID, Kitob_ID, Azo_ID, Olingan_sana, Qaytarish_sanasi)
VALUES (4, 4, 4, #6/5/2026#, #6/19/2026#);

INSERT INTO Ijara (ID, Kitob_ID, Azo_ID, Olingan_sana, Qaytarish_sanasi)
VALUES (5, 5, 5, #5/20/2026#, #6/3/2026#);

INSERT INTO Ijara (ID, Kitob_ID, Azo_ID, Olingan_sana, Qaytarish_sanasi)
VALUES (6, 1, 2, #6/10/2026#, #6/24/2026#);

INSERT INTO Ijara (ID, Kitob_ID, Azo_ID, Olingan_sana, Qaytarish_sanasi)
VALUES (7, 3, 1, #4/1/2025#, #4/15/2025#);

/* -- Javobgarlik (muddati o'tganlarga jarima) */
INSERT INTO Javobgarlik (ID, Ijara_ID, Kechiktirilganlik_uchun_jarima)
VALUES (1, 1, 25000);

INSERT INTO Javobgarlik (ID, Ijara_ID, Kechiktirilganlik_uchun_jarima)
VALUES (2, 2, 18000);

INSERT INTO Javobgarlik (ID, Ijara_ID, Kechiktirilganlik_uchun_jarima)
VALUES (3, 5, 31000);

INSERT INTO Javobgarlik (ID, Ijara_ID, Kechiktirilganlik_uchun_jarima)
VALUES (4, 7, 45000);


/* ================================================================
   2. SO'ROV: MUDDATI O'TGAN KITOBLAR RO'YXATI
   QUERY: Overdue Books List
   (Access da "Query" sifatida saqlang: qry_MuddatiOtgan)
   ================================================================ */

SELECT
    K.Nomi                                          AS [Kitob nomi],
    K.Muallif,
    A.Ism & " " & A.Familiya                        AS [Azoning ismi],
    A.Azo_raqami                                    AS [A'zo raqami],
    I.Olingan_sana                                  AS [Olingan sana],
    I.Qaytarish_sanasi                              AS [Qaytarish sanasi],
    DateDiff("d", I.Qaytarish_sanasi, Date())       AS [Kechikkan kunlar],
    IIf(IsNull(J.Kechiktirilganlik_uchun_jarima),
        0,
        J.Kechiktirilganlik_uchun_jarima)           AS [Jarima (so'm)]
FROM (Ijara AS I
    INNER JOIN Kitoblar AS K ON I.Kitob_ID = K.ID
    INNER JOIN Azolar   AS A ON I.Azo_ID   = A.ID)
    LEFT JOIN Javobgarlik AS J ON J.Ijara_ID = I.ID
WHERE I.Qaytarish_sanasi < Date()
ORDER BY DateDiff("d", I.Qaytarish_sanasi, Date()) DESC;


/* ================================================================
   3. HISOBOT: OYLIK IJARA HISOBOTI
   REPORT: Monthly Rental Summary
   (Access da "Query" sifatida saqlang: qry_OylikHisobot)
   ================================================================ */

SELECT
    Format(I.Olingan_sana, "yyyy-mm")          AS [Oy],
    Count(I.ID)                                AS [Jami ijaralar],
    Count(DISTINCT I.Azo_ID)                   AS [Faol a'zolar],
    Count(DISTINCT I.Kitob_ID)                 AS [Ijaradagi kitoblar],
    Sum(IIf(IsNull(J.Kechiktirilganlik_uchun_jarima),
            0,
            J.Kechiktirilganlik_uchun_jarima)) AS [Jami jarimalar (so'm)]
FROM Ijara AS I
    LEFT JOIN Javobgarlik AS J ON J.Ijara_ID = I.ID
GROUP BY Format(I.Olingan_sana, "yyyy-mm")
ORDER BY Format(I.Olingan_sana, "yyyy-mm") DESC;


/* ================================================================
   QOSHIMCHA SO'ROVLAR / BONUS QUERIES
   ================================================================ */

/* Eng ko'p o'qilgan kitoblar */
SELECT
    K.Nomi              AS [Kitob nomi],
    K.Muallif,
    Count(I.ID)         AS [Ijara soni]
FROM Kitoblar AS K
    LEFT JOIN Ijara AS I ON I.Kitob_ID = K.ID
GROUP BY K.Nomi, K.Muallif
ORDER BY Count(I.ID) DESC;


/* Hozirda mavjud nusxalar soni */
SELECT
    K.Nomi                                                  AS [Kitob nomi],
    K.Soni                                                  AS [Jami nusxa],
    K.Soni - Count(IIf(I.Qaytarish_sanasi >= Date(), 1, Null)) AS [Mavjud nusxa]
FROM Kitoblar AS K
    LEFT JOIN Ijara AS I ON I.Kitob_ID = K.ID
GROUP BY K.Nomi, K.Soni
ORDER BY K.Nomi;
