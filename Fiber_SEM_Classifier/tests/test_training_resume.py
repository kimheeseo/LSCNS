"""Notebook control-flow checks without Drive, CUDA, pretrained weights, or training."""
import ast
import contextlib
import io
import json
from pathlib import Path
from types import SimpleNamespace
import unittest

NOTEBOOK = Path(__file__).resolve().parents[1] / 'SEM_CNN_classifier_epoch3.ipynb'
CELLS = json.loads(NOTEBOOK.read_text())['cells']
SOURCE = ''.join(CELLS[5]['source'])
NAMES = {'training_device', 'checkpoint_resume_message', 'restore_training_scaler'}
module = ast.parse(SOURCE)
helpers = ast.Module(body=[node for node in module.body
                          if isinstance(node, ast.FunctionDef) and node.name in NAMES],
                     type_ignores=[])


class Scaler:
    def __init__(self, enabled):
        self.enabled = enabled
        self.loaded = []

    def is_enabled(self):
        return self.enabled

    def load_state_dict(self, value):
        if not value:
            raise RuntimeError('CUDA GradScaler cannot load a disabled CPU scaler state')
        self.loaded.append(value)


class ResumeTests(unittest.TestCase):
    def setUp(self):
        self.gpu = False
        torch = SimpleNamespace(
            device=lambda name: SimpleNamespace(type=name),
            cuda=SimpleNamespace(is_available=lambda: self.gpu),
        )
        self.scope = {'torch': torch}
        exec(compile(helpers, str(NOTEBOOK), 'exec'), self.scope)

    def test_cpu_is_rejected_before_expensive_training(self):
        with self.assertRaisesRegex(RuntimeError, 'GPU를 사용할 수 없습니다'):
            self.scope['training_device'](require_gpu=True)

    def test_explicit_cpu_execution_remains_available(self):
        output = io.StringIO()
        with contextlib.redirect_stdout(output):
            device = self.scope['training_device']()
        self.assertEqual(device.type, 'cpu')
        self.assertIn('CPU 학습', output.getvalue())

    def test_gpu_preflight_accepts_cuda(self):
        self.gpu = True
        self.assertEqual(self.scope['training_device'](require_gpu=True).type, 'cuda')

    def test_incomplete_run_resumes_at_next_epoch(self):
        self.assertIn('7번째 Epoch부터 재개', self.scope['checkpoint_resume_message'](6, 10, 2, 7))

    def test_early_stopped_run_does_not_claim_to_resume(self):
        message = self.scope['checkpoint_resume_message'](9, 10, 7, 7)
        self.assertIn('Early Stopping 완료', message)
        self.assertNotIn('재개', message)

    def test_completed_or_reduced_epoch_budget_goes_to_evaluation(self):
        for budget in [3, 10]:
            message = self.scope['checkpoint_resume_message'](10, budget, 1, 7)
            self.assertIn('추가 학습 없이', message)
            self.assertNotIn('11번째', message)

    def test_cpu_checkpoint_empty_scaler_can_move_to_gpu(self):
        scaler = Scaler(True)
        with contextlib.redirect_stdout(io.StringIO()):
            self.scope['restore_training_scaler'](scaler, {})
            self.scope['restore_training_scaler'](scaler, None)
        self.assertEqual(scaler.loaded, [])

    def test_gpu_scaler_history_is_preserved_on_gpu_resume(self):
        scaler = Scaler(True)
        state = {'scale': 65536., 'growth_factor': 2.}
        self.scope['restore_training_scaler'](scaler, state)
        self.assertEqual(scaler.loaded, [state])

    def test_gpu_checkpoint_can_move_to_disabled_cpu_scaler(self):
        scaler = Scaler(False)
        self.scope['restore_training_scaler'](scaler, {'scale': 65536.})
        self.assertEqual(scaler.loaded, [])

    def test_python_cells_compile(self):
        for i, cell in enumerate(CELLS):
            if cell['cell_type'] != 'code':
                continue
            source = ''.join(line for line in cell['source'] if not line.startswith(('%', '!')))
            compile(source, f'cell-{i}', 'exec')


if __name__ == '__main__':
    unittest.main()
