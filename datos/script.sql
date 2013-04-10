CREATE TABLE rubros (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  descripcion VARCHAR(50) NOT NULL
);

INSERT INTO rubros (id, descripcion) VALUES (1, 'Sin rubro');

CREATE TABLE productos (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  codigo_barras VARCHAR(50) NULL,
  descripcion  VARCHAR(50) NOT NULL,
  idrubro INTEGER NOT NULL DEFAULT 0,
  stock_minimo FLOAT NULL NOT NULL DEFAULT 0,
  stock FLOAT NULL NOT NULL DEFAULT 0,
  costo FLOAT NOT NULL DEFAULT 0,
  precio FLOAT NOT NULL DEFAULT 0
);

CREATE TABLE usuarios (
  usuario VARCHAR(8) NOT NULL,
  contrasena VARCHAR(8) NOT NULL,
  es_admin TINYINT(1) NOT NULL DEFAULT 0,
  UNIQUE (contrasena),
  PRIMARY KEY (usuario)
);

INSERT INTO usuarios (usuario, contrasena, es_admin) VALUES ('admin', 'admin', 1);
