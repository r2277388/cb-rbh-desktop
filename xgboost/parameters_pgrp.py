def pgrp_parameters():
	"""
    Each key is a tuple representing a pgrp and a flbl flag, and the value is a dictionary of xgboost parameters.
    
    Returns:
        dict: A dictionary where keys are tuples and values are dictionaries of xgboost parameters.
    """
	return {
        ('ART', 'B'): {'n_estimators': 400, 'max_depth': 3, 'learning_rate': 0.01, 'booster': 'gbtree', 'base_score': 0.5},
        ('ART', 'F'): {'n_estimators': 400, 'max_depth': 3, 'learning_rate': 0.01, 'booster': 'gbtree', 'base_score': 0.1},
        ('CHL', 'B'): {'n_estimators': 400, 'max_depth': 3, 'learning_rate': 0.01, 'booster': 'gbtree', 'base_score': 0.5},
        ('CHL', 'F'): {'n_estimators': 400, 'max_depth': 3, 'learning_rate': 0.01, 'booster': 'gbtree', 'base_score': 0.1},
        ('CPA', 'B'): {'n_estimators': 400, 'max_depth': 3, 'learning_rate': 0.01, 'booster': 'gbtree', 'base_score': 0.1},
        ('CPA', 'F'): {'n_estimators': 400, 'max_depth': 3, 'learning_rate': 0.01, 'booster': 'gbtree', 'base_score': 0.1},
        ('ENT', 'B'): {'n_estimators': 400, 'max_depth': 3, 'learning_rate': 0.01, 'booster': 'gbtree', 'base_score': 0.5},
        ('ENT', 'F'): {'n_estimators': 400, 'max_depth': 3, 'learning_rate': 0.01, 'booster': 'gbtree', 'base_score': 0.1},
        ('FWN', 'B'): {'n_estimators': 400, 'max_depth': 3, 'learning_rate': 0.01, 'booster': 'gbtree', 'base_score': 0.1},
        ('FWN', 'F'): {'n_estimators': 400, 'max_depth': 3, 'learning_rate': 0.01, 'booster': 'gbtree', 'base_score': 0.1},
        ('LIF', 'B'): {'n_estimators': 400, 'max_depth': 3, 'learning_rate': 0.01, 'booster': 'gbtree', 'base_score': 0.1},
        ('LIF', 'F'): {'n_estimators': 400, 'max_depth': 3, 'learning_rate': 0.01, 'booster': 'gbtree', 'base_score': 0.1},
        ('BAR', 'B'): {'n_estimators': 400, 'max_depth': 3, 'learning_rate': 0.01, 'booster': 'gbtree', 'base_score': 0.5},
        ('BAR', 'F'): {'n_estimators': 400, 'max_depth': 9, 'learning_rate': 0.01, 'booster': 'gbtree', 'base_score': 0.5},
        ('CPB', 'B'): {'n_estimators': 400, 'max_depth': 3, 'learning_rate': 0.01, 'booster': 'gbtree', 'base_score': 0.1},
        ('CPB', 'F'): {'n_estimators': 400, 'max_depth': 3, 'learning_rate': 0.01, 'booster': 'gbtree', 'base_score': 0.1},
        ('CCB', 'B'): {'n_estimators': 400, 'max_depth': 3, 'learning_rate': 0.01, 'booster': 'gbtree', 'base_score': 0.1},
        ('CCB', 'F'): {'n_estimators': 400, 'max_depth': 6, 'learning_rate': 0.01, 'booster': 'gbtree', 'base_score': 0.5},
        ('RID', 'B'): {'n_estimators': 400, 'max_depth': 3, 'learning_rate': 0.01, 'booster': 'gbtree', 'base_score': 0.1},
        ('RID', 'F'): {'n_estimators': 400, 'max_depth': 6, 'learning_rate': 0.01, 'booster': 'gbtree', 'base_score': 0.1},
        ('GAM', 'B'): {'n_estimators': 400, 'max_depth': 3, 'learning_rate': 0.01, 'booster': 'gbtree', 'base_score': 0.1},
        ('GAM', 'F'): {'n_estimators': 400, 'max_depth': 3, 'learning_rate': 0.01, 'booster': 'gbtree', 'base_score': 0.1},
        ('PTC', 'B'): {'n_estimators': 400, 'max_depth': 3, 'learning_rate': 0.01, 'booster': 'gbtree', 'base_score': 0.5},
        ('PTC', 'F'): {'n_estimators': 400, 'max_depth': 9, 'learning_rate': 0.01, 'booster': 'gbtree', 'base_score': 0.5}
        }