package Lab2;


import org.apache.commons.math3.distribution.TDistribution;
import org.apache.commons.math3.stat.correlation.Covariance;
import org.apache.commons.math3.stat.descriptive.DescriptiveStatistics;
import org.apache.commons.math3.stat.descriptive.SummaryStatistics;
import org.apache.commons.math3.stat.descriptive.moment.Variance;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class StatisticsCalculator {

    public static double calculateGeometricMean(List<Double> sample) {
        DescriptiveStatistics stats = new DescriptiveStatistics();
        sample.forEach(stats::addValue);
        return stats.getGeometricMean();
    }

    public static double calculateArithmeticMean(List<Double> sample) {
        DescriptiveStatistics stats = new DescriptiveStatistics();
        sample.forEach(stats::addValue);
        return stats.getMean();
    }

    public static double calculateStandardDeviation(List<Double> sample) {
        DescriptiveStatistics stats = new DescriptiveStatistics();
        sample.forEach(stats::addValue);
        return stats.getStandardDeviation();
    }

    public static double calculateRange(List<Double> sample) {
        DescriptiveStatistics stats = new DescriptiveStatistics();
        sample.forEach(stats::addValue);
        return stats.getMax() - stats.getMin();
    }

    public static double calculateCovariance(List<Double> sample1, List<Double> sample2) {
        double[] array1 = sample1.stream()
                .mapToDouble(Double::doubleValue)
                .toArray();
        double[] array2 = sample2.stream()
                .mapToDouble(Double::doubleValue)
                .toArray();
        return new Covariance().covariance(array1, array2);

    }

    public static int countElements(List<Double> sample) {
        return sample.size();
    }

    public static double calculateCoefficientOfVariation(List<Double> sample) {
        DescriptiveStatistics stats = new DescriptiveStatistics();
        sample.forEach(stats::addValue);
        return stats.getStandardDeviation() / stats.getMean();
    }

    public static double[] calculateConfidenceInterval(List<Double> sample, double confidenceLevel) {
        SummaryStatistics stats = new SummaryStatistics();
        sample.forEach(stats::addValue);
        TDistribution tDistribution =new TDistribution(stats.getN() - 1);
        double standardError = stats.getStandardDeviation() / Math.sqrt(stats.getN());
        double marginOfError = tDistribution.inverseCumulativeProbability(1.0 - (1.0 - confidenceLevel) / 2) * standardError;
        double[] interval = new double[2];
        interval[0] = stats.getMean() - marginOfError;
        interval[1] = stats.getMean() + marginOfError;
        return interval;
    }

    public static double calculateVariance(List<Double> sample) {
        Variance variance = new Variance();
        return variance.evaluate(sample.stream().mapToDouble(Double::doubleValue).toArray());
    }

    public static Map<String, Double> calculateMinMax(List<Double> sample) {
        Map<String, Double> minMax = new HashMap<>();
        DescriptiveStatistics stats = new DescriptiveStatistics();
        sample.forEach(stats::addValue);
        minMax.put("Min", stats.getMin());
        minMax.put("Max", stats.getMax());
        return minMax;
    }
}

