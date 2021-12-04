package net.summer.fastexcel.xlsx.reader;

class StringSupplier implements Supplier {
    private final java.util.function.Supplier<String> supplier;

    StringSupplier(final String val) {
        this.supplier = () -> val;
    }

    @Override
    public Object getContent() {
        return this.supplier.get();
    }
}
