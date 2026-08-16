# Release Notes v0.7.1

This patch fixes a workflow loop found in live testing: after `ネイル 風水 就活` was approved and promoted to a Creator case, the parent `ネイル・ペディキュア 風水` case still appeared as pending and `▶ 次に進む` returned to SERP review.

v0.7.1 propagates resolution state from the derived Creator case back to the parent validation workflow. It also repairs already-promoted v0.7.0 cases without requiring the user to repeat the promotion.

The separate issue where a new Creator case currently carries an existing article URL in `対象記事` is intentionally not changed in this patch; it will be verified at the SBM handoff boundary before changing the inter-product contract.
