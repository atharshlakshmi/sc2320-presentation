const PptxGenJS = require('pptxgenjs');
const pptx = new PptxGenJS();

pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'Monotone Classification — Lower Bounds';

const TOTAL = 8;

function makeSlide(num) {
  const sl = pptx.addSlide();
  sl.background = { color: 'FFFFFF' };
  sl.addText(`${num} / ${TOTAL}`, {
    x: 11.5, y: 7.1, w: 1.5, h: 0.3,
    fontSize: 11, color: 'BBBBBB', align: 'right',
  });
  return sl;
}

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 1
// ─────────────────────────────────────────────────────────────────────────────
{
  const sl = makeSlide(1);
  sl.addText([
    { text: 'Can we do better?  ', options: { fontSize: 40, bold: true, color: '111111' } },
  ], { x: 1, y: 1.2, w: 11, h: 0.8 });

  sl.addText([
    { text: 'RPE gives ', options: { color: '222222' } },
    { text: 'expected error at most 2k*', options: { color: '222222', highlight: 'FFE066' } },
    { text: ' in O(w Log(n/w)) probes', options: { color: '222222' } },
  ], { x: 1, y: 2.5, w: 11, h: 0.5, fontSize: 22 });

  sl.addText([
    { text: 'Coreset gives ', options: { color: '222222' } },
    { text: 'error at most (1+e)k*', options: { color: '222222', highlight: 'FFE066' } },
    { text: ' in O(w/e^2 Log(n/w) log n) probes', options: { color: '222222' } },
  ], { x: 1, y: 3.2, w: 11, h: 0.5, fontSize: 22 });

  sl.addText('Are these the best we can do?', {
    x: 1, y: 4.5, w: 11, h: 0.5, fontSize: 22, color: '888888',
  });

  sl.addNotes(`So far in these lecture notes we have seen two algorithms for the monotone classification problem. The first one, RPE, which stands for Random Probe and Eliminate, gives us a classifier whose expected error is at most two times the optimal monotone error, k-star. And it does this using O of w times Log of n over w probes, where w is the dominance width of the partially ordered set. The second algorithm, which is based on relative-comparison coresets, goes further: for any epsilon you choose, it gives you a classifier with error at most one plus epsilon times k-star, and it uses O of w over epsilon squared times Log of n over w times log n probes.

Now, both of these results are interesting, but they naturally raise the question: can we do better? Is there some smarter algorithm out there that could achieve the same accuracy with fewer probes? Or, maybe, could we get a tighter approximation ratio without paying more in terms of probes?

This is exactly what this section of the paper addresses. The authors prove three lower bound results that essentially say no, you cannot do significantly better. These are not just arguments about specific algorithms. These are information-theoretic lower bounds, meaning they apply to every possible algorithm, no matter how clever. By the end of this section, we will see that RPE and the coreset algorithm are, up to small logarithmic factors, the best we can hope for. Let us start with the most extreme case.`);
}

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 2
// ─────────────────────────────────────────────────────────────────────────────
{
  const sl = makeSlide(2);
  sl.addText([
    { text: 'Finding k* exactly is hopeless', options: { fontSize: 40, bold: true, color: '111111' } },
  ], { x: 1, y: 1.2, w: 11, h: 0.8 });

  sl.addText([
    { text: 'Any algorithm that outputs the ', options: { color: '222222' } },
    { text: 'optimal classifier', options: { color: '222222', highlight: 'FFE066' } },
  ], { x: 1, y: 2.8, w: 11, h: 0.5, fontSize: 24 });

  sl.addText([
    { text: 'must probe ', options: { color: '222222' } },
    { text: 'Omega(n)', options: { color: '222222', highlight: 'FFE066' } },
    { text: ' elements', options: { color: '222222' } },
  ], { x: 1, y: 3.4, w: 11, h: 0.5, fontSize: 24 });

  sl.addText('even in 1D, even if k* is known in advance', {
    x: 1, y: 4.5, w: 11, h: 0.5, fontSize: 22, color: '888888',
  });

  sl.addNotes(`The first lower bound is Theorem 10, and it deals with the case where epsilon equals zero, meaning we want the exact optimal monotone classifier, not an approximation. The result says that any algorithm which outputs a classifier with error exactly equal to k-star, with probability greater than two-thirds, must probe at least omega of n elements in expectation. And this holds even in the simplest possible setting: one-dimensional data, and even if the algorithm is told the value of k-star upfront.

Why is this significant? Well, if your dataset has n elements, probing omega of n elements is essentially the same as reading the entire dataset. So this theorem tells us that there is no shortcut. If you insist on exact optimality, you are forced to look at almost everything.

This is a strong negative result. It means that the problem of finding the exact best monotone classifier is fundamentally different from the approximation version. No matter how smart your algorithm is, whether it is deterministic or randomized, adaptive or non-adaptive, it cannot avoid reading a constant fraction of the input.

To prove this, the authors construct a specific family of hard inputs where the exact optimum is impossible to find without essentially inspecting every element. Let me show you what this family looks like.`);
}

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 3
// ─────────────────────────────────────────────────────────────────────────────
{
  const sl = makeSlide(3);
  sl.addText('How do we prove it?', {
    x: 1, y: 0.8, w: 11, h: 0.8, fontSize: 40, bold: true, color: '111111',
  });

  sl.addText([
    { text: 'Construct a ', options: { color: '222222' } },
    { text: 'hard input family', options: { color: '222222', highlight: 'FFE066' } },
    { text: ' \u2014 n/2 pairs, one anomaly hidden', options: { color: '222222' } },
  ], { x: 1, y: 1.9, w: 11, h: 0.5, fontSize: 22 });

  // Draw pairs
  const pairY = 3.0;
  const pairDefs = [
    { x: 1.0,  anom: false },
    { x: 2.4,  anom: false },
    { x: 3.8,  dots: true  },
    { x: 4.8,  anom: true  },
    { x: 6.4,  dots: true  },
    { x: 7.4,  anom: false },
  ];

  for (const p of pairDefs) {
    if (p.dots) {
      sl.addText('\u2026', { x: p.x, y: pairY, w: 0.6, h: 1.2, fontSize: 28, color: 'BBBBBB', align: 'center', valign: 'middle' });
      continue;
    }
    const bc = p.anom ? 'E67700' : 'DDDDDD';
    const bg = p.anom ? 'FFF8E1' : 'FFFFFF';
    sl.addShape(pptx.ShapeType.roundRect, {
      x: p.x, y: pairY, w: 1.15, h: 1.35,
      fill: { color: bg }, line: { color: bc, width: 2 }, rectRadius: 0.1,
    });
    const topC = p.anom ? 'E67700' : '3B5BDB';
    sl.addShape(pptx.ShapeType.ellipse, { x: p.x+0.35, y: pairY+0.15, w: 0.35, h: 0.35, fill: { color: topC }, line: { color: bg } });
    sl.addText(p.anom ? '\u2212' : '+', { x: p.x+0.35, y: pairY+0.15, w: 0.35, h: 0.35, fontSize: 14, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' });
    const botC = p.anom ? 'E67700' : 'E03131';
    sl.addShape(pptx.ShapeType.ellipse, { x: p.x+0.35, y: pairY+0.55, w: 0.35, h: 0.35, fill: { color: botC }, line: { color: bg } });
    sl.addText('\u2212', { x: p.x+0.35, y: pairY+0.55, w: 0.35, h: 0.35, fontSize: 14, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' });
    sl.addText(p.anom ? 'anomaly' : 'normal', { x: p.x, y: pairY+1.0, w: 1.15, h: 0.3, fontSize: 11, color: p.anom ? 'E67700' : '999999', align: 'center' });
  }

  sl.addText([
    { text: '\u2192 algorithm must ', options: { color: '222222' } },
    { text: 'find the anomaly', options: { color: '222222', highlight: 'FFE066' } },
    { text: ' to be optimal', options: { color: '222222' } },
  ], { x: 1, y: 5.0, w: 11, h: 0.5, fontSize: 22 });

  sl.addNotes(`The proof constructs what we call a hard input family. Here is how it works. Take n elements and group them into n over 2 consecutive pairs: pair one is elements one and two, pair two is elements three and four, and so on. In every normal pair, the odd-numbered element gets label plus one and the even-numbered element gets label minus one.

Now, here is the twist. Exactly one of these pairs is an anomaly. In the anomaly pair, both elements get the same label, either both minus one or both plus one. So overall, we have n different possible inputs: n over 2 possible anomaly positions, each with two possible polarities.

The key observation is that k-star equals n over 2 minus 1 for every single input in this family. Why? Because any monotone classifier on a totally ordered set is just a threshold: everything up to some point gets minus one, everything after gets plus one. And for every normal pair, one element will be misclassified no matter where you put the threshold. So the minimum error is exactly one per normal pair, except at the anomaly where you can get both right.

But the optimal classifier is different depending on where the anomaly is and what its labels are. So the algorithm must figure out which pair is anomalous. And the only way to determine that is to probe elements in different pairs until you find the one that behaves differently. Since the anomaly could be anywhere, this requires searching through a linear number of pairs. That gives us the omega of n lower bound.`);
}

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 4
// ─────────────────────────────────────────────────────────────────────────────
{
  const sl = makeSlide(4);
  sl.addText([
    { text: "Yao\u2019s minimax principle", options: { fontSize: 40, bold: true, color: '111111' } },
  ], { x: 1, y: 1.2, w: 11, h: 0.8 });

  sl.addText([
    { text: 'Hard for every ', options: { color: '222222' } },
    { text: 'deterministic', options: { color: '222222', highlight: 'FFE066' } },
    { text: ' algorithm on average', options: { color: '222222' } },
  ], { x: 1, y: 2.8, w: 11, h: 0.5, fontSize: 24 });

  sl.addText([
    { text: '\u2192 hard for every ', options: { color: '222222' } },
    { text: 'randomized', options: { color: '222222', highlight: 'FFE066' } },
    { text: ' algorithm in the worst case', options: { color: '222222' } },
  ], { x: 1, y: 3.5, w: 11, h: 0.5, fontSize: 24 });

  sl.addText('family-cost >= n^2(1-c^2)/4  \u2192  at least one input needs Omega(n) probes', {
    x: 1, y: 5.0, w: 11, h: 0.5, fontSize: 20, color: '888888',
  });

  sl.addNotes(`Now I should explain how the proof actually goes through formally, because what I just described gives the intuition, but we need to handle the fact that algorithms can be randomized.

The formal proof uses a beautiful technique called Yao's minimax principle. The idea is this: instead of directly reasoning about randomized algorithms, which are hard to analyze because they make random choices, we instead reason about deterministic algorithms run against a randomly chosen input.

Here is the argument. First, we give the algorithm a boost: when it probes either element of a pair, it learns both labels for free. This only makes the algorithm stronger, so any lower bound we prove in this model also applies to the original model.

Now consider a deterministic algorithm that probes pairs in some fixed order. We use a counting argument, this is Lemma 11 in the paper. If the algorithm only probes t pairs, then for the n over 2 minus t pairs it never checked, it cannot distinguish between the two anomaly types. So it must get at least one of the two versions wrong for each unchecked pair. This gives us a lower bound on the total number of errors across the whole family.

At the same time, we can compute the total probing cost across the family. Combining these two, we get that the total family cost, the sum of probes over all inputs, is at least n squared times one minus c squared over 4. Since there are n inputs in the family, at least one of them must require omega of n probes.

Yao's principle then tells us: if every deterministic algorithm has high average cost on this distribution, then every randomized algorithm has high worst-case cost. And that completes the proof of Theorem 10.`);
}

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 5
// ─────────────────────────────────────────────────────────────────────────────
{
  const sl = makeSlide(5);
  sl.addText([
    { text: 'Even approximation is expensive', options: { fontSize: 40, bold: true, color: '111111' } },
  ], { x: 1, y: 1.2, w: 11, h: 0.8 });

  sl.addText([
    { text: 'Any ', options: { color: '222222' } },
    { text: 'c-approximate', options: { color: '222222', highlight: 'FFE066' } },
    { text: ' algorithm needs', options: { color: '222222' } },
  ], { x: 1, y: 2.8, w: 11, h: 0.5, fontSize: 24 });

  sl.addText('Omega(w * Log(n / k*w))', {
    x: 1, y: 3.5, w: 11, h: 0.7, fontSize: 36, bold: true, color: '111111',
  });

  sl.addText('probes', { x: 1, y: 4.1, w: 11, h: 0.5, fontSize: 24, color: '222222' });

  sl.addText('\u2192 RPE is essentially optimal for constant-factor approximation', {
    x: 1, y: 5.2, w: 11, h: 0.5, fontSize: 20, color: '888888',
  });

  sl.addNotes(`So the exact case requires omega of n probes. What if we relax our requirements and allow some approximation? Theorem 13 answers this for constant-factor approximation.

The result says: for any constant c greater than or equal to 1, any algorithm that guarantees expected error at most c times k-star must use at least omega of w-prime times Log of n-prime over w-prime probes on some input of width w and optimal error k-star.

The proof uses a clever reduction to the realizable case. The idea is to take a hard family where k-star equals zero, the realizable setting, and then add so-called dummy points to artificially create a nonzero k-star. These dummy points are carefully placed so that each real point is surrounded by c times k dummy points that have the same correct label.

Why does this work? Because if the algorithm misclassifies even one real point, the c-k surrounding dummies will also be misclassified, giving a total error of at least c times k plus 1. Since the algorithm promises error at most c times k-star equals c times k, it cannot afford to get any real point wrong. This means the algorithm is forced to correctly classify every real point, which is exactly the realizable problem on the original points.

And we already know from a prior result that the realizable problem requires omega of w times Log of n over w probes. So the algorithm inherits that lower bound.

The bottom line is that RPE, which achieves O of w times Log of n over w probes, is essentially optimal for constant-factor approximation. You cannot beat it by more than a small factor depending on k-star.`);
}

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 6
// ─────────────────────────────────────────────────────────────────────────────
{
  const sl = makeSlide(6);
  sl.addText([
    { text: 'The 1/e^2 barrier', options: { fontSize: 40, bold: true, color: '111111' } },
  ], { x: 1, y: 1.2, w: 11, h: 0.8 });

  sl.addText([
    { text: 'Any ', options: { color: '222222' } },
    { text: '(1+e)-approximate', options: { color: '222222', highlight: 'FFE066' } },
    { text: ' algorithm needs', options: { color: '222222' } },
  ], { x: 1, y: 2.8, w: 11, h: 0.5, fontSize: 24 });

  sl.addText('Omega(w / e^2)', {
    x: 1, y: 3.5, w: 11, h: 0.7, fontSize: 36, bold: true, color: '111111',
  });

  sl.addText('probes', { x: 1, y: 4.1, w: 11, h: 0.5, fontSize: 24, color: '222222' });

  sl.addText('\u2192 like telling apart two biased coins \u2014 you need Theta(1/e^2) flips', {
    x: 1, y: 5.2, w: 11, h: 0.5, fontSize: 20, color: '888888',
  });

  sl.addNotes(`Now we come to the most important lower bound: Theorem 14. This one deals with arbitrary epsilon, and it says that any algorithm guaranteeing expected error at most one plus epsilon times k-star must use at least omega of w over epsilon squared probes. This holds even for two-dimensional data.

This is a very strong result because it tells us that the one over epsilon squared dependence in the coreset algorithm is not an artifact of the algorithm's design. It is a fundamental barrier. No algorithm can avoid it.

The hard input construction is different from the previous ones. Here we place n over w points at each of w pairwise incomparable locations. Think of these locations as forming an antichain, no location dominates any other. Then each location independently gets a random bias: either mu-1 which equals one half minus gamma over 2, or mu-2 which equals one half plus gamma over 2, where gamma is about 9 epsilon. Each point at a location gets label plus one with probability equal to that location's bias, independently.

The crux of the difficulty is a coin-distinguishing problem. At each location, the algorithm sees a sequence of random labels, some plus one, some minus one. The distribution of these labels is either slightly biased toward minus one or slightly biased toward plus one. The difference between the two biases is tiny, it is proportional to epsilon. And the classical result from statistics, related to Le Cam's method, says that you need theta of one over epsilon squared samples to reliably tell these two distributions apart.

Since the algorithm must figure out the bias at each of the w locations to classify correctly, it needs at least omega of w over epsilon squared probes in total.`);
}

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 7
// ─────────────────────────────────────────────────────────────────────────────
{
  const sl = makeSlide(7);
  sl.addText('Why Omega(w/e^2)?', {
    x: 1, y: 0.8, w: 11, h: 0.8, fontSize: 40, bold: true, color: '111111',
  });

  sl.addText([
    { text: 'Each location is a ', options: { color: '222222' } },
    { text: 'biased coin', options: { color: '222222', highlight: 'FFE066' } },
    { text: ' \u2014 flip it enough to tell which side it favours', options: { color: '222222' } },
  ], { x: 1, y: 1.9, w: 11, h: 0.5, fontSize: 22 });

  // Coin boxes
  sl.addShape(pptx.ShapeType.roundRect, {
    x: 1.5, y: 2.9, w: 3.5, h: 1.6,
    fill: { color: 'FFFFFF' }, line: { color: 'DDDDDD', width: 2 }, rectRadius: 0.12,
  });
  sl.addText('\u00BD \u2212 \u03B5', { x: 1.5, y: 3.1, w: 3.5, h: 0.6, fontSize: 30, bold: true, color: '111111', align: 'center' });
  sl.addText('slightly more \u22121', { x: 1.5, y: 3.75, w: 3.5, h: 0.4, fontSize: 14, color: '999999', align: 'center' });

  sl.addText('vs', { x: 5.4, y: 3.2, w: 0.8, h: 0.6, fontSize: 28, color: 'CCCCCC', align: 'center' });

  sl.addShape(pptx.ShapeType.roundRect, {
    x: 6.5, y: 2.9, w: 3.5, h: 1.6,
    fill: { color: 'FFFFFF' }, line: { color: 'DDDDDD', width: 2 }, rectRadius: 0.12,
  });
  sl.addText('\u00BD + \u03B5', { x: 6.5, y: 3.1, w: 3.5, h: 0.6, fontSize: 30, bold: true, color: '111111', align: 'center' });
  sl.addText('slightly more +1', { x: 6.5, y: 3.75, w: 3.5, h: 0.4, fontSize: 14, color: '999999', align: 'center' });

  sl.addText([
    { text: '\u2192 need ', options: { color: '222222' } },
    { text: 'Theta(1/e^2) samples per location', options: { color: '222222', highlight: 'FFE066' } },
    { text: '  \u00D7  w locations  =  Omega(w/e^2)', options: { color: '222222' } },
  ], { x: 1, y: 5.2, w: 11, h: 0.5, fontSize: 22 });

  sl.addNotes(`Let me give you the intuition for why the proof works, because I think this is the most elegant part.

Imagine you have a coin, and you want to figure out whether it lands heads with probability one half minus epsilon, or one half plus epsilon. These two coins are very similar. The difference in their bias is only 2 epsilon. If you only flip the coin a few times, you will see roughly the same mix of heads and tails either way. It is only when you flip it many times, specifically on the order of one over epsilon squared times, that the law of large numbers kicks in and you can reliably tell the two coins apart.

This is a classical result in statistics. It follows from what is known as Le Cam's method, or in the context of this paper, the Anthony-Bartlett distinguishing lemma. The authors formalize this by defining two random processes: RP-1, where all labels are generated upfront, and RP-2, where labels are generated on demand as the algorithm probes. They prove that these two processes are equivalent in terms of the expected approximation ratio, this is Lemma 15.

The advantage of RP-2 is that we can track exactly how many times the algorithm probes each location. If the total probe budget is less than w times M over 8, where M is theta of one over epsilon squared, then by a simple averaging argument, more than half the locations have been probed fewer than M times. At these under-probed or light locations, Lemma 19 shows the algorithm guesses the bias wrong with probability greater than one quarter.

Each wrong guess contributes a certain amount of excess error. Summing over all the light locations, the expected approximation ratio exceeds one plus epsilon, violating the algorithm's guarantee. This contradiction proves that the algorithm must use at least omega of w over epsilon squared probes.`);
}

// ─────────────────────────────────────────────────────────────────────────────
// SLIDE 8
// ─────────────────────────────────────────────────────────────────────────────
{
  const sl = makeSlide(8);
  sl.addText([
    { text: 'The full picture', options: { fontSize: 40, bold: true, color: '111111' } },
  ], { x: 1, y: 0.7, w: 11, h: 0.8 });

  // Table
  const colX = [1.0, 2.8, 5.8, 8.8];
  const colW = [1.8, 3.0, 3.0, 2.0];
  const headers = ['Regime', 'Upper bound', 'Lower bound', 'Gap'];
  for (let c = 0; c < 4; c++) {
    sl.addText(headers[c], { x: colX[c], y: 1.8, w: colW[c], h: 0.4, fontSize: 13, bold: true, color: '888888' });
  }
  sl.addShape(pptx.ShapeType.line, { x: 1, y: 2.2, w: 9.8, h: 0, line: { color: 'EEEEEE', width: 1 } });

  const rows = [
    ['e = 0',       'O(n)',                          'Omega(n)',              { t: 'tight',   c: '0CA678' }],
    ['constant c',  'O(w Log(n/w))',                 'Omega(w Log(n/k*w))',   { t: 'tight*',  c: '0CA678' }],
    ['any e > 0',   'O(w/e^2 Log log)',              'Omega(w/e^2)',          { t: 'polylog', c: 'E67700' }],
  ];

  for (let r = 0; r < rows.length; r++) {
    const y = 2.4 + r * 0.65;
    for (let c = 0; c < 4; c++) {
      const cell = rows[r][c];
      if (typeof cell === 'string') {
        sl.addText(cell, { x: colX[c], y, w: colW[c], h: 0.5, fontSize: 18, color: '222222' });
      } else {
        sl.addText(cell.t, { x: colX[c], y, w: colW[c], h: 0.5, fontSize: 18, bold: true, color: cell.c });
      }
    }
    sl.addShape(pptx.ShapeType.line, { x: 1, y: y + 0.55, w: 9.8, h: 0, line: { color: 'F5F5F5', width: 1 } });
  }

  sl.addText([
    { text: '\u2192 our algorithms are ', options: { color: '222222' } },
    { text: 'near-optimal', options: { color: '222222', highlight: 'FFE066' } },
  ], { x: 1, y: 5.0, w: 11, h: 0.5, fontSize: 24 });

  sl.addNotes(`Let me now put everything together. This table summarizes the three regimes and how the upper and lower bounds compare.

In the first regime, epsilon equals zero, we want the exact optimum. The upper bound is trivially O of n because you can just read everything. And Theorem 10 says the lower bound is omega of n. So these are tight. Finding the exact optimum is genuinely as hard as reading the entire input.

In the second regime, we allow a constant approximation factor c. The RPE algorithm gives an upper bound of O of w times Log of n over w. Theorem 13 gives a lower bound of omega of w times Log of n over k-star-w. When k-star is not too large, specifically when k-star is at most n over w to the power one minus delta for some small delta, these two bounds match. So RPE is optimal for constant-factor approximation in the typical case.

In the third regime, we want a one-plus-epsilon approximation for any epsilon. The coreset algorithm achieves O of w over epsilon squared times Log of n over w times log n probes. Theorem 14 gives a lower bound of omega of w over epsilon squared. So there is a gap, but it is only polylogarithmic, just some Log and log n factors.

The conclusion is remarkable: our algorithms are near-optimal. The dominant terms, w in the cost, one over epsilon squared in the sampling, are all necessary. The only remaining open question is whether the polylogarithmic gap can be closed, either by improving the upper bound or by strengthening the lower bound. But the core complexity of the monotone classification problem is now fully understood: it is theta of w over epsilon squared, up to polylogarithmic factors. And that is a very satisfying conclusion.`);
}

// ── Write ────────────────────────────────────────────────────────────────────
pptx.writeFile({ fileName: 'C:/Users/kex03/OneDrive/Desktop/sc2320-presentation/Chunk4_LowerBounds_Minimal.pptx' })
  .then(() => console.log('Done: Chunk4_LowerBounds_Light.pptx'))
  .catch(err => console.error('Error:', err));
