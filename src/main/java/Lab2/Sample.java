package Lab2;


import org.apache.commons.math3.distribution.TDistribution;
import org.apache.commons.math3.stat.correlation.Covariance;
import org.apache.commons.math3.stat.descriptive.DescriptiveStatistics;
import org.apache.commons.math3.stat.descriptive.SummaryStatistics;
import org.apache.commons.math3.stat.descriptive.moment.Variance;


import java.util.List;


public class Sample {
    final List<Double> sample;
    final int id;
    final private DescriptiveStatistics stats;

    public Sample(List<Double> sample, int id) {
        this.sample = sample;
        this.id=id;
        stats = new DescriptiveStatistics();
        sample.forEach(stats::addValue);
    }

    private double calculateGeometricMean() {
        return stats.getGeometricMean();
    }

    private double calculateArithmeticMean() {
        return stats.getMean();
    }

    private double calculateStandardDeviation() {
        return stats.getStandardDeviation();
    }

    private double calculateRange() {
        return stats.getMax() - stats.getMin();
    }

    double calculateCovariance(List<Double> sample2) {
        double[] array1 = sample.stream()
                .mapToDouble(Double::doubleValue)
                .toArray();
        double[] array2 = sample2.stream()
                .mapToDouble(Double::doubleValue)
                .toArray();
        return new Covariance().covariance(array1, array2);

    }

    private int countElements() {
        return sample.size();
    }

    private double calculateCoefficientOfVariation() {
        return stats.getStandardDeviation() / stats.getMean();
    }

    private double[] calculateConfidenceInterval(double confidenceLevel) {
        SummaryStatistics stats = new SummaryStatistics();
        sample.forEach(stats::addValue);
        TDistribution tDistribution = new TDistribution(stats.getN() - 1);
        double standardError = stats.getStandardDeviation() / Math.sqrt(stats.getN());
        double marginOfError = tDistribution.inverseCumulativeProbability(1.0 - (1.0 - confidenceLevel) / 2) * standardError;
        double[] interval = new double[2];
        interval[0] = stats.getMean() - marginOfError;
        interval[1] = stats.getMean() + marginOfError;
        return interval;
    }

    private double calculateVariance() {
        Variance variance = new Variance();
        return variance.evaluate(sample.stream().mapToDouble(Double::doubleValue).toArray());
    }

    private double[] calculateMinMax() {
        return new double[]{
                stats.getMin(),
                stats.getMax()
        };
    }

    public double[] getResult() {
        return new double[]{
                calculateGeometricMean(),
                calculateArithmeticMean(),
                calculateStandardDeviation(),
                calculateRange(),
                countElements(),
                calculateCoefficientOfVariation(),
                calculateConfidenceInterval(0.95)[0],
                calculateConfidenceInterval(0.95)[1],
                calculateVariance(),
                calculateMinMax()[0],
                calculateMinMax()[1]
        };
    }

}

