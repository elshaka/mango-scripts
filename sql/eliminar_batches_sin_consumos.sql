DELETE FROM batches WHERE id NOT IN (SELECT batch_id FROM batch_hoppers_lots);
