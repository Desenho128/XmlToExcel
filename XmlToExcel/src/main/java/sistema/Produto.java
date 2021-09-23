package sistema;

import java.math.BigDecimal;

class Produto{
		
		private String nfNumero;
		private String cod;
		private String descricao;
		private String ncm;
		private BigDecimal valorDeCompra;
		private Integer quant;
		private String ean;

		
		public Produto() {

		}

		public Produto(String cod, String descricao, Integer quant) {
			this.cod = cod;
			this.descricao = descricao;
			this.quant = quant;
		}
				
		public Produto(String cod, String descricao, Integer quant, String ean) {
			super();
			this.cod = cod;
			this.descricao = descricao;
			this.quant = quant;
			this.ean = ean;
		}
		
		public Produto(String nfNumero, String cod, String descricao, String ncm, BigDecimal valorDeCompra, Integer quant,
				String ean) {
			super();
			this.nfNumero = nfNumero;
			this.cod = cod;
			this.descricao = descricao;
			this.ncm = ncm;
			this.valorDeCompra = valorDeCompra;
			this.quant = quant;
			this.ean = ean;
		}

		public String getCod() {
			return cod;
		}

		public void setCod(String cod) {
			this.cod = cod;
		}

		public String getDescricao() {
			return descricao;
		}

		public void setDescricao(String descricao) {
			this.descricao = descricao;
		}

		public Integer getQuant() {
			return quant;
		}

		public void setQuant(Integer quant) {
			this.quant = quant;
		}

		public String getEan() {
			return ean;
		}

		public void setEan(String ean) {
			this.ean = ean;
		}
		
		
		public String getNfNumero() {
			return nfNumero;
		}

		public void setNfNumero(String nfNumero) {
			this.nfNumero = nfNumero;
		}

		public String getNcm() {
			return ncm;
		}

		public void setNcm(String ncm) {
			this.ncm = ncm;
		}

		public BigDecimal getValorDeCompra() {
			return valorDeCompra;
		}

		public void setValorDeCompra(BigDecimal valorDeCompra) {
			this.valorDeCompra = valorDeCompra;
		}

		@Override
		public String toString() {
			return new StringBuilder().append("COD: ").append(this.getCod()).append(" Descrição: ").append(this.getDescricao()).append(" Quant: ").append(this.getQuant()).append(" EAN: ").append(this.getEan()).toString();
		}
	 }