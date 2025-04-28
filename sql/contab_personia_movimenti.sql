CREATE TABLE `contab_personia_movimenti` (
  `Id` int NOT NULL PRIMARY KEY AUTO_INCREMENT,
  `Descrizione` varchar(255) CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci DEFAULT NULL COMMENT 'Descrizione del movimento contabile ed eventuale giustificazione della spesa',
  `Codice` int DEFAULT NULL COMMENT 'Collegamento alla lista dei codici del piano dei conti',
  `Data` date DEFAULT NULL COMMENT 'Data del movimento contabile',
  `Montante` decimal(10,2) DEFAULT NULL COMMENT 'Se negativo, dunque corrisponde a delle uscite o spese viene chiamato Avere. Se positivo, dunque corrisponde a delle entrate o guadagni viene chiamato Dare.'
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci COMMENT='Movimenti contabili gestiti come libro/giornale aziendale di contabilità ordinaria';
