XGBoost model to VB function
============================

XGBoost is a popular tree-based gradient boosting library (but you knew that..).

Often we want to create prototypes of models for expert users to assess before deployment occurs.  Excel remains a favourite due to ubiquity, user familiarity and ease of creating pretty dashboards.  VBA is the most practical way of executing a pre-trained model in Excel.  If that's what you're trying to do you can do worse than the code here.

It started from https://github.com/Shiutang-Li/xgboost_deployment/blob/master/deploy_xgboost.ipynb . However, aside from needing to generate VBA a few improvements were made:
* Multi-class (not just binary) classification is handled.
* A bug is fixed in the regular expression for grabbing numbers in scientific notation.
* The function prototype is explicit in the names of parameters, which makes it easier when calling from Excel.

On the other hand, I didn't implement missing value treatment as I did not need it.  It could be added easily, however.
