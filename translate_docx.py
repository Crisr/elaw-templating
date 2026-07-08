import json
import os


def load_config(path="config.json"):
    with open(path) as f:
        return json.load(f)


def get_provider(config, name=None):
    if name is None:
        name = config["default_provider"]
    try:
        provider = config["providers"][name].copy()
    except KeyError:
        raise ValueError(f"Provider '{name}' not found in config. Available: {list(config['providers'].keys())}")
    api_key_env = provider.pop("api_key_env", None)
    if api_key_env:
        provider["api_key"] = os.environ.get(api_key_env)
        if not provider["api_key"]:
            from dotenv import load_dotenv
            load_dotenv()
            provider["api_key"] = os.environ.get(api_key_env)
    return provider


def test_config_loading():
    import json, tempfile, os
    cfg = {
        "default_provider": "test",
        "providers": {
            "test": {
                "base_url": "http://test/v1",
                "model": "test-model",
                "api_key_env": None
            }
        }
    }
    with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
        json.dump(cfg, f)
        f.flush()
        result = load_config(f.name)
        assert result == cfg
        os.unlink(f.name)


def test_get_provider():
    cfg = {
        "default_provider": "test",
        "providers": {
            "test": {
                "base_url": "http://test/v1",
                "model": "test-model",
                "api_key_env": None
            }
        }
    }
    provider = get_provider(cfg, "test")
    assert provider["base_url"] == "http://test/v1"
    assert provider["model"] == "test-model"


if __name__ == "__main__":
    test_config_loading()
    test_get_provider()
    print("PASS")
